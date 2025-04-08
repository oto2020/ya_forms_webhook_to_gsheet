const express = require('express');
const bodyParser = require('body-parser');

const app = express();
const PORT = 12345;

// Чтобы парсить JSON в теле запроса
app.use(bodyParser.json());

// Принимаем POST-запросы с формы
app.post('/webhook', (req, res) => {
  console.log('Получены данные от Яндекс Формы:');
  console.log(req.body);

  // Можно потом сохранить или отправить куда-то ещё

  res.status(200).send('Данные получены');
});

app.listen(PORT, () => {
  console.log(`Сервер запущен: http://localhost:${PORT}/webhook`);
});
