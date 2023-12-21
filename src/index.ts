const REGION = '関東';
const REGIONS: { [key: string]: string } = {
  北海道: 'hokkaido',
  東北: 'tohoku',
  関東: 'kanto',
  '甲信越・北陸': 'koshinetsu',
  東海: 'tokai',
  近畿: 'kinki',
  '中国・四国': 'chugoku',
  九州: 'kyushu',
};
const ORIGIN = 'http://www.sej.co.jp';
const NEW_ITEM_DIR = '/products/a/thisweek/area/';
const NEW_ITEM_URL = ORIGIN + NEW_ITEM_DIR;
const QUERY = '/1/l100/';

const SEVEN_COLOR = [
  '#f58220', // セブンオレンジ
  '#00a54f', // セブングリーン
  '#ee1c23', // セブンレッド
];

interface Attachment {
  color?: string;
  title: string;
  title_link: string;
  thumb_url: string;
  fields: Array<{ title: string; value: string; short: boolean }>;
}

function getMatch(input: string, regex: RegExp): string | null {
  const match = input.match(regex);
  return match ? match[1] : null;
}

function main(): void {
  const attachments: Attachment[] = [];

  const html = UrlFetchApp.fetch(NEW_ITEM_URL + REGIONS[REGION] + QUERY).getContentText();
  const items = html.match(/<div class="list_inner[^>]*>[\s\S]*?<\/div>\s*<\/div>/g) || [];

  if (!items) {
    return;
  }

  for (let i = 0; i < items.length; ++i) {
    const itemDetails = extractItemDetails(items[i]);
    if (itemDetails) {
      attachments.push({
        color: SEVEN_COLOR[i % SEVEN_COLOR.length],
        ...itemDetails,
      });
    }
  }

  // スプレッドシートに書き込み
  writeToSpreadsheet(attachments);
}

function extractField(
  itemHtml: string,
  regex: RegExp,
  fieldTitle: string,
): { title: string; value: string; short: boolean } {
  const value = getMatch(itemHtml, regex) || '';
  return { title: fieldTitle, value: value, short: true };
}

function extractItemDetails(itemHtml: string): Attachment {
  const link = getMatch(itemHtml, /<a href="([^"]+)"/) || '';
  const image = getMatch(itemHtml, /data-original="([^"]+)"/) || '';
  const name =
    getMatch(itemHtml, /<div class="item_ttl"><p><a href="[^"]+">([^<]+)<\/a><\/p><\/div>/) || '';

  return {
    title_link: ORIGIN + link,
    thumb_url: image,
    title: name,
    fields: [
      extractField(itemHtml, /<div class="item_price"><p>([^<]+)<\/p><\/div>/, '値段'),
      extractField(itemHtml, /<div class="item_launch"><p>([^<]+)<\/p><\/div>/, '販売時期'),
      extractField(
        itemHtml,
        /<div class="item_region"><p><span>販売地域：<\/span>([^<]+)<\/p><\/div>/,
        '販売地域',
      ),
    ],
  };
}

function writeToSpreadsheet(data: Attachment[]): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('シート1'); // スプレッドシートのシート名を適宜設定

  if (!sheet) {
    throw new Error('シートがありません.');
  }

  // 既存のデータをクリア
  sheet.clear();

  // ヘッダーの追加
  sheet.appendRow(['商品名', '値段', '販売時期', '販売地域', 'リンク', '画像URL']);

  // 新しいデータを書き込み
  data.forEach((item) => {
    const row = [
      item.title,
      item.fields.find((f) => f.title === '値段')?.value || '',
      item.fields.find((f) => f.title === '販売時期')?.value || '',
      item.fields.find((f) => f.title === '販売地域')?.value || '',
      item.title_link,
      item.thumb_url,
    ];
    sheet.appendRow(row);
  });
}
