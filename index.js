import axios from 'axios';
import * as cheerio from 'cheerio';

const TEAMS_WEBHOOK_URL = process.env.TEAMS_WEBHOOK_URL;
const LONGBLACK_URL = 'https://www.longblack.co';

async function getLongBlackArticle() {
  try {
    // 1. 롱블랙 메인 페이지 HTML 가져오기
    const { data } = await axios.get(LONGBLACK_URL, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
      }
    });
    
    const $ = cheerio.load(data);

    // 2. 오늘의 아티클 데이터 추출 (롱블랙의 클래스명/구조에 맞게 셀렉터 수정이 필요할 수 있습니다)
    // 아래 셀렉터는 예시이며, 실제 페이지 구조를 확인 후 맞춰주셔야 합니다.
    const title = $('.today-note-title').text().trim() || '오늘의 롱블랙 아티클';
    const description = $('.today-note-paragraph').text().trim() || '오늘이 지나면 읽을 수 없는 롱블랙의 인사이트를 확인하세요.';
    const articleLink = $('.today-note-link').attr('href') || '';
    const fullLink = articleLink.startsWith('http') ? articleLink : `${LONGBLACK_URL}${articleLink}`;

    return { title, description, fullLink };
  } catch (error) {
    console.error('크롤링 중 에러 발생:', error.message);
    throw error;
  }
}

async function sendToTeams(article) {
  if (!TEAMS_WEBHOOK_URL) {
    console.error('TEAMS_WEBHOOK_URL 세팅이 되어있지 않습니다.');
    return;
  }

  // MS 팀즈 전용 Adaptive Card 포맷 템플릿
  const payload = {
    type: "message",
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: {
          type: "AdaptiveCard",
          version: "1.4",
          body: [
            {
              type: "TextBlock",
              text: "☕️ 오늘의 롱블랙 도착",
              weight: "Bolder",
              size: "Large",
              color: "Accent"
            },
            {
              type: "TextBlock",
              text: article.title,
              weight: "Bolder",
              size: "Medium",
              wrap: true
            },
            {
              type: "TextBlock",
              text: article.description,
              wrap: true,
              italic: true
            }
          ],
          actions: [
            {
              type: "Action.OpenUrl",
              title: "아티클 읽으러 가기",
              url: article.fullLink
            }
          ],
          $schema: "http://adaptivecards.io/schemas/adaptive-card.json"
        }
      }
    ]
  };

  try {
    await axios.post(TEAMS_WEBHOOK_URL, payload);
    console.log('팀즈로 메시지 전송 성공!');
  } catch (error) {
    console.error('팀즈 전송 중 에러 발생:', error.response?.data || error.message);
  }
}

async function run() {
  const article = await getLongBlackArticle();
  await sendToTeams(article);
}

run();
