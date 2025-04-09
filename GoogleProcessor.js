
const { google } = require('googleapis');
class GoogleProcessor {
  // Статические свойства для хранения объектов
  static keys = require('./Keys.json');
  static sid;
  static gid;
  static client = new google.auth.JWT(this.keys.client_email, null, this.keys.private_key, ['https://www.googleapis.com/auth/spreadsheets']);
  static gsapi = google.sheets({ version: 'v4', auth: this.client });

  // Метод для инициализации аутентификации
  static async init(SID, GID) {
    this.sid = SID;
    this.gid = GID;
    return new Promise((resolve, reject) => {
      this.client.authorize((error, tokens) => {
        if (error) {
          console.error('Authorization error:', error);
          reject(error);
        } else {
          console.log('Connected...');
          resolve(tokens);
        }
      });
    });
  }


  // ---------- ОСНОВНЫЕ ФУНКЦИИ ------

  // 🔍 Получение имени листа по gid
  static async getSheetNameByGid() {
    try {
      const response = await this.gsapi.spreadsheets.get({
        spreadsheetId: this.sid,
      });

      const sheets = response.data.sheets;
      const sheet = sheets.find(s => s.properties.sheetId.toString() === this.gid.toString());

      if (!sheet) {
        throw new Error(`Sheet with gid ${this.gid} not found`);
      }

      return sheet.properties.title;
    } catch (error) {
      console.error('Error getting sheet name by gid:', error);
      throw error;
    }
  }

  // Функция для удаления всех строк, где в столбце 'A' присутствует указанное значение
  static async deleteRowsByColumnValue(columnValue) {
    try {
      let sheetName = await this.getSheetNameByGid();
      // Читаем все данные из указанного столбца 'A'
      const range = `${sheetName}!A:A`;
      let values = await this.readRange(range);

      // Находим индексы строк, где значение совпадает с columnValue
      const rowsToDelete = [];
      for (let i = 0; i < values.length; i++) {
        if (values[i][0] === columnValue) {
          rowsToDelete.push(i + 1); // Строки нумеруются с 1
        }
      }

      if (rowsToDelete.length === 0) {
        console.log('No rows found with the specified value.');
        return;
      }

      // Удаляем строки, начиная с последней, чтобы не нарушить индексацию
      for (const row of rowsToDelete.reverse()) {
        await this.deleteRow(row);
        console.log(`Deleted row ${row}`);
      }

    } catch (error) {
      console.error('Error deleting rows:', error);
      // throw error;
    }
  }

  // Вспомогательная функция для удаления строки
  static async deleteRow(row) {
    try {
      const batchUpdateRequest = {
        spreadsheetId: this.sid,
        resource: {
          requests: [
            {
              deleteDimension: {
                range: {
                  sheetId: this.gid,
                  dimension: 'ROWS',
                  startIndex: row - 1,
                  endIndex: row
                }
              }
            }
          ]
        }
      };

      await this.gsapi.spreadsheets.batchUpdate(batchUpdateRequest);
    } catch (error) {
      console.error('Error deleting row:', error);
      throw error;
    }
  }

 
  // Определение номера первой пустой строки
  static async findFirstEmptyRow() {
    try {
      let sheetName = await this.getSheetNameByGid();
      // Читаем все данные из листа
      const range = `${sheetName}`; // Читаем весь лист
      const values = await this.readRange(range);

      // Найдем первую пустую строку
      let rowNumber = values.length + 1; // Начинаем с предположения, что пустая строка будет после последней строки

      // Проходим через все строки и ищем первую пустую
      for (let i = 0; i < values.length; i++) {
        if (values[i].every(cell => cell === undefined || cell === '')) {
          rowNumber = i + 1; // Строки нумеруются с 1
          break;
        }
      }

      return rowNumber;
    } catch (error) {
      console.error('Error finding the first empty row:', error);
      throw error;
    }
  }

  // Метод для получения заголовков листа
  static async getHeadersRow(rowNumber) {
    try {
      let sheetName = await this.getSheetNameByGid();
      // Читаем первую строку (заголовки) из указанного листа
      const range = `${sheetName}!${rowNumber}:${rowNumber}`; // Заданная строка
      const values = await this.readRange(range);

      // Проверяем, есть ли заголовки
      if (values.length === 0) {
        console.log('No headers found.');
        return [];
      }

      // Возвращаем первую строку как заголовки
      const headers = values[0]; // Первый элемент - это заголовки
      return headers;
    } catch (error) {
      console.error('Error getting headers:', error);
      return null;
    }
  }


  // находит первое встречающееся в столбце А название подразделения 
  static async findRowInFirstColumn(currentDivision) {
    let sheetName = await this.getSheetNameByGid();
    let divisionColumnLetter = 'A';
    let values = await this.readRange(`${sheetName}!${divisionColumnLetter}:${divisionColumnLetter}`);
    values = values.map(el => el[0]);
    // console.log(values);
    // let currentDivision = 'Групповые программы';
    let rowNumber = values.findIndex(el => el === currentDivision) + 1; // в какой строке искомое подразделение
    // console.log(rowNumber);
    return rowNumber;
  }
  // находит второй встречающийся месяц формата "2024-08" и выдает номер строки
  static async columnNumberInRow(rowNumber, foundingValue) {
    try {
      let sheetName = await this.getSheetNameByGid();
      let values = await this.readRange(`${sheetName}!${rowNumber}:${rowNumber}`);
      values = values[0]; // так как массив двумерный
  
      let colIndex = values.findIndex(el => el === foundingValue);
  
      if (colIndex === -1) {
        return null; // если не найдено
      }
  
      return colIndex + 1;
    } catch (error) {
      return null;
    }
  }

  static async columnNumbersBySearchValuesInRow(rowNumber, searchValues = []) {
    try {
      if (!Array.isArray(searchValues) || searchValues.length === 0) return [];
  
      const sheetName = await this.getSheetNameByGid();
      const values = await this.readRange(`${sheetName}!${rowNumber}:${rowNumber}`);
      const row = values[0]; // первая (и единственная) строка
  
      const result = [];
  
      for (const value of searchValues) {
        const colIndex = row.findIndex(cell => cell === value);
        if (colIndex !== -1) {
          result.push({ value, columnNumber: colIndex + 1 });
        }
      }
  
      return result;
    } catch (error) {
      console.error('Ошибка при поиске колонок по значениям:', error);
      return [];
    }
  }
  
  

  static async getCellValue(cellAddress) {
    const options = {
      spreadsheetId: this.sid,
      range: cellAddress,
    };
  
    try {
      const resp = await this.gsapi.spreadsheets.values.get(options);
      const value = resp.data.values?.[0]?.[0] ?? null;
      return value;
    } catch (error) {
      console.error(`Ошибка при получении значения из ${cellAddress}:`, error.message);
      return null;
    }
  }

  
  // Возвращает год-месяц в формате 2024-09
  static getCurrentYearMonth() {
    const date = new Date();
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Добавляем 1, т.к. месяцы начинаются с 0
    return `${year}-${month}`;
  }

  // Функция для преобразования номера столбца в букву столбца
  static getColumnLetter(columnNumber) {
    let letter = '';
    let temp = columnNumber;

    while (temp > 0) {
      let mod = (temp - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      temp = Math.floor((temp - mod) / 26);
    }
    return letter;
  }

  // ЧТЕНИЕ
  static async readRange(range) {
    const opt = {
      spreadsheetId: this.sid,
      range: range
    };
    let dataObtained = await this.gsapi.spreadsheets.values.get(opt);
    return dataObtained.data.values;
  }

// ЗАПИСЬ
static async writeRange(range, values) {
  try {
    const updateOptions = {
      spreadsheetId: this.sid,
      range: range,
      valueInputOption: 'USER_ENTERED',
      resource: { values: values }
    };
    const resp = await this.gsapi.spreadsheets.values.update(updateOptions);
    return resp;
  } catch (error) {
    console.error('❌ Ошибка при записи в Google Таблицу:', error.message);
    return null;
  }
}


  // 
  static async writeCell(cellAdress, value) {
    const updateOptions = {
      spreadsheetId: this.sid,
      range: cellAdress,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [[value]] }
    };
    try {
      const resp = await this.gsapi.spreadsheets.values.update(updateOptions);
      if (resp.status === 200) {
        return true;
      } else {
        console.error('Unexpected response status:', resp.status);
        return false;
      }
    } catch (error) {
      console.error(`Failed to write to cell ${cellAdress}: ${value}`, error.message);
      return false;
    }
  }
  

  static async getSheetsList() {
    try {
      const response = await this.gsapi.spreadsheets.get({
        spreadsheetId: this.sid,
      });

      const sheets = response.data.sheets;
      const sheetNames = sheets.map(sheet => sheet.properties.title);

      console.log('Sheet names:', sheetNames);
      return sheetNames;
    } catch (error) {
      console.error('Error getting sheet names:', error);
    }
  }


}



module.exports = GoogleProcessor;
