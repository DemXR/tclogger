# Ведение журнала автотестов
## Описание
Модуль позволяет вести журнал автотестов в Excel документе, и в перспективе в любых иных форматах. Главная особенность: возможность сохранения скриншотов экрана во время ведения журнала.

## Инструкция
### Ведение журнала в Excel
Создайте журнал, указав директорию в которую будут сохранены все необходимые для журнала файлы:
```
from tclogger import create_xlsx_logger
logger = create_xlsx_logger(directory="./tests/")
```
Используйте методы info, success, warning, error для ведения записей с соответствующей классификацей:
```
logger.info(case_name="Кейс #1", message="Информационный текст")
logger.success(case_name="Кейс #1")
logger.warning(case_name="Кейс #2", message="Текст предупреждения")
logger.error(case_name="Кейс #3", message="Текст несоответствия", make_screenshot=True)
logger.save(open_file=True)
```
Аргументы методов:

**case_name** - Наименование тестового кейса;
**message** - Текст сообщения;
**make_screenshot** - Нужно ли создать скриншот.
