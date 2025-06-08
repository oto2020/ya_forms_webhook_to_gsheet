const express = require('express');
const bodyParser = require('body-parser');
const GoogleProcessor = require('./GoogleProcessor'); // путь к твоему GoogleProcessor
require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');

const app = express();
const PORT = process.env.PORT;

app.use(bodyParser.json());

bot = new TelegramBot(process.env.TELEGRAM_TOKEN, { polling: false });

// Форматирует данные из поля, возвращает одну строку
function valueStringFromAnswerData(answerData, header) {
  // console.log(header);
  // console.log(answerData);
  let qv = answerData[header];
  // получаем тип ответа
  let answerType = qv?.question?.answer_type?.slug;

  // если значения выбора
  if (answerType == 'answer_choices') {
    return qv.value.map(v => v.text).join('\n');
  }

  // если файл(ы)
  else if (answerType == 'answer_files') {
    return qv.value.map(v => v.name).join('\n');
  }

  // если булево
  else if (answerType == 'answer_boolean') {
    return qv.value ? 'Да' : 'Нет';
  }

  // если дата (yyyy-MM-dd) переведем в dd.MM.yyyy
  else if (answerType == 'answer_date') {
    const [year, month, day] = qv.value.split('-');
    return `${day}.${month}.${year}`;
  }

  // В остальных случаях возвращаем значение строки
  let valueString = qv.value;

  // Но перед этим поработаем с номером телефона, удалив лишние пробелы и дефисы из него
  if (header.includes('phone') || valueString.startsWith('+')) {
    valueString = normalizePhoneNumber(valueString);
  }

  return valueString;
}

function normalizePhoneNumber(phone) {
  if (typeof phone === 'string' && phone.startsWith('+')) {
    return '\'' + phone.replace(/[\s-]/g, '');
  }
  return phone;
}

function convertToGoogleSheetsDateTime(isoString) {
  const date = new Date(isoString);
  if (isNaN(date)) return '';

  // Получаем строку в формате "ГГГГ-ММ-ДД ЧЧ:ММ:СС"
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0'); // Месяцы от 0
  const day = String(date.getDate()).padStart(2, '0');

  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');

  return `${day}.${month}.${year} ${hours}:${minutes}`;
}

function expandAnswerArray(answerArray) {
  // Находим максимальный номер колонки
  const maxColumnNumber = Math.max(...answerArray.map(a => a.columnNumber));

  // Создаем массив значений длиной в maxColumnNumber и заполняем пустыми строками
  const rowValues = new Array(maxColumnNumber).fill('');

  // Вставляем значения в соответствующие места
  for (const item of answerArray) {
    rowValues[item.columnNumber - 1] = item.value;
  }
  return rowValues;
}

// удаляет форматирование телеграм
function removeTelegramFormatting(text) {
  return text
    // Заменить переносы строк на пробелы
    .replace(/[\r\n]+/g, ' ')
    // MarkdownV2 спецсимволы
    .replace(/[*_~`[\]()><#+\-=|{}.!\\]/g, ' ')
    // Блок кода (```код```)
    .replace(/```[^```]*```/g, ' ')
    // Моноширинный текст (`код`)
    .replace(/`[^`]*`/g, ' ')
    // Markdown-ссылки [текст](url)
    .replace(/\[.*?\]\(.*?\)/g, ' ')
    // HTML-теги <b>, <i>, <code>, <a> и прочие
    .replace(/<\/?[^>]+(>|$)/g, ' ')
    // Сжать множественные пробелы
    .replace(/\s+/g, ' ')
    .trim();
}

app.post('/webhook', async (req, res) => {
  console.log('Получены данные от Яндекс Формы:');

  try {
    const base64Encoded = req.body?.params?.answer;
    if (!base64Encoded) throw new Error('Нет данных в params.answer');

    // 1. Декодируем из base64 в строку
    const jsonString = Buffer.from(base64Encoded, 'base64').toString('utf-8');
    // console.log('Декодированная строка:', jsonString.slice(0, 200)); // ограничим вывод

    // 2. Парсим JSON
    const parsed = JSON.parse(jsonString);
    // console.log('parsed:', parsed);

    // 3. Работаем по прежней логике
    let createdAt = parsed.created;
    let answerData = parsed.answer?.data;



    const headers = Object.keys(answerData);
    let sid, gid, tgGroupId;
    for (let header of headers) {
      let valueString = valueStringFromAnswerData(answerData, header);

      console.log(`header: ${header}\nvalueString: ${valueString}\n`);

      // если попались gid и sid
      if (header == 'sid') {
        sid = valueString;
      }
      if (header == 'gid') {
        gid = valueString;
      }
      if (header == 'tgGroupId') {
        tgGroupId = valueString;
      }
    }


    // РАБОТА С ГУГЛ-ЧАСТЬЮ
    if (gid && sid) {
      try {
        await GoogleProcessor.init(sid, gid);
        const sheetName = await GoogleProcessor.getSheetNameByGid();
        console.log('Имя листа:', sheetName);

        // найдем первую свободную строку
        let rowNumber = await GoogleProcessor.findFirstEmptyRow();

        const sheetLink = `https://docs.google.com/spreadsheets/d/${sid}/edit#gid=${gid}&range=A${rowNumber}`;
        console.log(`\nПолучен ответ из яндекс.форм, записываю в строку: ${rowNumber} документа:\n ${sheetLink}`);

        // Собираем Массив с овтетами
        let answerArray = [];

        // заполним дату создания заявки. она в первом столбце
        let columnNumber = 0;
        try {
          columnNumber = await GoogleProcessor.columnNumberInRow(process.env.HEADER_ROW_NUMBER, 'createdAt');
          if (columnNumber) {
            createdAt = convertToGoogleSheetsDateTime(createdAt);
            answerArray.push({ header: 'createdAt', headerRus: 'Дата заполнения', columnNumber, value: createdAt });
          }
          console.log({ columnNumber });
        } catch (e) {
          console.error('Пиздец, не могу найти createdAt. Вот ошибка:\n', e)
        }




        // ищем ячейки с латиноязычными заголовками
        let headersColumnNumbers = await GoogleProcessor.columnNumbersBySearchValuesInRow(process.env.HEADER_ROW_NUMBER, headers);
        let headerRusArray = await GoogleProcessor.getHeadersRow(process.env.HEADERRUS_ROW_NUMBER);
        // console.table(headersColumnNumbers);
        // console.table(headerRusArray);

        // Наполняем answerArray значениями полей с номерами строк и заголовками
        for (let headerColumnNumber of headersColumnNumbers) {
          let columnNumber = headerColumnNumber.columnNumber;
          let header = headerColumnNumber.value;
          let value = valueStringFromAnswerData(answerData, header);
          answerArray.push({
            header,                                               // заголовок eng
            headerRus: headerRusArray[columnNumber - 1],          // заголовок rus
            columnNumber,                                         // номер столбца
            value                                                 // значение из формы
          });
        }

        // Собранный массив объектов с ответами
        console.table(answerArray);


        // отправим сообщение в телеграм
        if (tgGroupId) {
          try {
            let messageText =
              `Ответ с формы *${sheetName}*:\n\n` +
              answerArray.map(a => `*${a.headerRus}:* ${ removeTelegramFormatting(a.value)}`).join('\n') +
              `\n\n[Открыть таблицу](${sheetLink})`;
            await bot.sendMessage(tgGroupId, messageText, {
              parse_mode: 'Markdown'
            });
          } catch (e) {
            bot.sendMessage(tgGroupId, `Получен с формы *${sheetName}*:\n\n...\n\n[Открыть таблицу](${sheetLink})`);
          }
        }

        // Значения, которые будут записаны на лист sheetName в строку rowNumber
        let valuesRow = expandAnswerArray(answerArray);
        console.table(valuesRow);

        // Запишем строку в гугл-таблицу
        if (await GoogleProcessor.writeRange(`${sheetName}!${rowNumber}:${rowNumber}`, [valuesRow])) {
          console.log(`Значения записаны в строку ${rowNumber}`);
        }

      } catch (e) {
        console.error('Не удалось обработать ответ: отправить сообщение и заполнить гугл-таблицу...\n' + e);
      }
    }

    res.status(200).send('Данные получены');
  } catch (e) {
    console.error('Ошибка при парсинге:', e.message);
    res.status(400).send('Ошибка парсинга JSON');
  }
});






app.listen(PORT, () => {
  console.log(`Сервер запущен: http://localhost:${PORT}/webhook`);
});
