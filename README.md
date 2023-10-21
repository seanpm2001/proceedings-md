# Proceedings

Этот репозиторий содержит скрипт для автоматической конвертации `.md` в `.docx`,
соответствующий требованиям форматирования Трудов Института системного
программирования РАН.

## Prerequisites:

```bash
git clone https://gitlab.ispras.ru/build-race-condition-detection/proceedings
cd proceedings
npm install
sudo apt-get install pandoc
```

## Conversion:

Файл `sample.md` содержит стандартный шаблон статьи для Трудов ИСП РАН,
представленный в `.md`-формате. Для генерации `.docx`-документа потребуется:

```
cd sample
node ../src/main.js sample.md sample.docx
````

## Notes

Скрипт был разработан в рекордные сроки, и ещё сырой.
Ошибки мало обрабатываются и могут быть нечитаемыми. MR приветствуются!

Некоторые версии ворда ругаются на то, что документ повреждён, но все равно
открывают его. OpenOffice падает даже на документе из `pandoc` без модификаций.
Перед отправкой документа рекомендуется прогнать его через обычный ворд, перепроверить
форматирование, и пересохранить.

Happy researching!