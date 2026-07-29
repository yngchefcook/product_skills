---
name: presentation-code-interpreter
description: Create and quality-check polished PPTX presentations inside Code Interpreter using the uploaded portable builder bundle. Use for executive decks, proposals, reports, product and strategy narratives, educational presentations, investor-style materials, data-driven slides, and requests that require generated or bundled visuals without a company-specific template.
---

# Presentation Code Interpreter

Создавать цельные профессиональные презентации: сильная логика, естественный
текст, крупная типографика, нативные диаграммы, качественные изображения и
проверенная вёрстка. Использовать загруженные `builder.py`, `helpers.py`,
`image_filenames.py` и ресурсы; не переписывать движок для каждой презентации.

## 1. Найти и подключить bundle

Работать только внутри каталога загруженных файлов, обычно `/mnt/data` или
текущей рабочей директории. Не искать рекурсивно по всему диску.

Если загружен ZIP с файлами Code Interpreter:

```python
from pathlib import Path
import zipfile

data_root = Path("/mnt/data") if Path("/mnt/data").exists() else Path.cwd()
archives = list(data_root.glob("*code-interpreter*.zip"))
if not archives:
    archives = list(data_root.glob("*.zip"))
if not archives:
    raise FileNotFoundError("Не найден ZIP с presentation bundle")

archive_path = None
for candidate in archives:
    with zipfile.ZipFile(candidate) as archive:
        if any(Path(name).name == "builder.py" for name in archive.namelist()):
            archive_path = candidate
            break
if archive_path is None:
    raise FileNotFoundError("В загруженных ZIP не найден builder.py")

bundle_dir = data_root / "presentation_bundle"
bundle_dir.mkdir(parents=True, exist_ok=True)
with zipfile.ZipFile(archive_path) as archive:
    archive.extractall(bundle_dir)
```

Найти `builder.py` только в ожидаемых уровнях:

```python
from pathlib import Path

roots = [Path.cwd()]
if Path("/mnt/data").exists():
    roots.insert(0, Path("/mnt/data"))

candidates = []
for root in roots:
    candidates.extend(root.glob("builder.py"))
    candidates.extend(root.glob("*/builder.py"))
    candidates.extend(root.glob("*/*/builder.py"))

if not candidates:
    raise FileNotFoundError("Не найден builder.py")
bundle_dir = candidates[0].resolve().parent
```

Подключить bundle и проверить среду:

```python
import sys

sys.path.insert(0, str(bundle_dir))

from builder import build_from_spec, lint_spec
from helpers import environment_report
from image_filenames import IMAGE_FILES, list_images

report = environment_report()
print(report)
if not report["ready"]:
    raise RuntimeError("Bundle или зависимости загружены не полностью")
```

Если отсутствуют зависимости, установить их из `requirements.txt` доступным
пакетным менеджером. Не переустанавливать уже работающую среду.

## 2. Определить коммуникационную задачу

До выбора слайдов сформулировать:

> К концу презентации [аудитория] должна [понять / выбрать / одобрить /
> сделать], потому что [главный вывод].

Определить аудиторию, цель, центральный тезис, необходимые доказательства,
ограничения по данным и желаемое действие. Если вводных достаточно, не
останавливаться ради подтверждения плана. Задать вопрос только когда отсутствует
решение, которое существенно меняет результат.

Выбрать одну сюжетную арку:

- контекст → ставки → доказательства → вывод → действие;
- вопрос → анализ → ответ;
- проблема → причины / варианты → рекомендация;
- текущее состояние → изменение → будущее состояние;
- процесс, хронология или обучающая прогрессия.

Повестка не является сюжетом. Каждый слайд должен создавать потребность в
следующем.

## 3. Написать содержание

- Давать каждому слайду одну работу и один главный вывод.
- Писать заголовок как мысль, которую можно произнести вслух, а не как название
  темы.
- Писать для аудитории, не показывать инструкции модели, план производства,
  внутренние заметки или технические комментарии.
- Использовать прямой естественный язык без рекламных клише и повторяющихся
  формул.
- Сокращать текст или менять композицию до уменьшения шрифта.
- Не создавать вымышленные факты, клиентов, цитаты, источники и результаты.
- Использовать сценарные цифры только по разрешению пользователя и подписывать
  каждый затронутый слайд: `Иллюстративный сценарий`, `Допущение` или
  `Пример, не фактические данные`.
- Указывать `source` для внешних фактов, графиков и таблиц.
- Не заканчивать пустым «Спасибо». Завершать решением, действием, синтезом или
  продуктивным вопросом.

Если пользователь не указал объём, использовать 7–10 слайдов. Увеличивать
количество только когда это необходимо для логики, а не для заполнения.

## 4. Выбрать визуальную систему

Прочитать `DESIGN_GUIDE.md`. Использовать одну тему на всю презентацию:

- `midnight`: технологии, финансы, стратегия;
- `paper`: аналитика, отчёты, обучение;
- `warm`: продуктовые и человеческие истории;
- `forest`: устойчивость, инфраструктура, долгосрочные программы.

Сохранять минимум:

- 50 pt для названия презентации;
- 35 pt для заголовков слайдов;
- 24 pt для подзаголовков и callout-заголовков;
- 16 pt для основного текста.

Предпочитать одну композицию вместо сетки UI-карточек. Избегать pills, badges,
псевдокнопок, вкладок, перегруженных dashboard-композиций и случайного декора.
Чередовать силуэты: визуальный якорь → аналитика → данные → statement.

## 5. Выбрать или создать изображения

Для готовой библиотеки:

```python
from image_filenames import list_images

print(list_images(keyword="AI"))
print(list_images(recommended_use="title"))
```

Передавать ID изображения в поле `image`, например `ai_neural_orbit`. Не
повторять одно изображение на соседних слайдах.

Если нужен новый визуал, использовать доступный инструмент генерации до сборки:

- текущий встроенный инструмент: `image_gen.imagegen`;
- совместимый MCP прежних runtime: `generate-image`;
- не импортировать эти инструменты в Python.

Для `image_gen.imagegen` сохранить выбранный результат в рабочую директорию и
передать локальный путь в `image`. Для `generate-image` передать возвращённый
HTTP(S) URL в `image_url`. Builder скачает, проверит и встроит его в PPTX.

Генерировать один самостоятельный сюжет одним вызовом. Для обложки и фона
использовать 16:9; для бокового визуала — 4:3 или 1:1. Заранее просить negative
space со стороны текста. Не создавать логотипы, названия организаций, водяные
знаки, читаемый синтетический текст, фальшивые интерфейсы или изображения,
которые выглядят как реальные данные.

Не рисовать растровые иллюстрации Python-кодом и не подменять визуал
декоративными фигурами. Подробный workflow находится в `IMAGE_GENERATION.md`.

## 6. Собрать spec

Перед сложной сборкой прочитать `SPEC_REFERENCE.md`. Доступны типы:

`title`, `section`, `statement`, `bullets`, `image`, `metrics`, `chart`,
`comparison`, `timeline`, `process`, `quote`, `table`, `summary`, `closing`.

Использовать нативный `chart`; не вставлять matplotlib, seaborn, plotly или
скриншоты графиков.

Минимальный пример:

```python
spec = {
    "filename": "executive_brief.pptx",
    "theme": "midnight",
    "slides": [
        {
            "type": "title",
            "title": "Главный вывод задаёт направление всей презентации",
            "subtitle": "Короткий контекст",
            "image": "ai_neural_orbit",
        },
        {
            "type": "statement",
            "title": "Один сильный тезис меняет рамку обсуждения",
            "body": "Короткое доказательство или следствие.",
        },
        {
            "type": "summary",
            "title": "Решение готово к переходу в план",
            "takeaway": "Одна синтезирующая мысль.",
            "actions": ["Назначить владельца", "Зафиксировать срок", "Начать"],
        },
        {
            "type": "closing",
            "title": "Следующий шаг превращает вывод в действие",
            "subtitle": "Владелец · срок · ожидаемый результат",
            "image": "closing_horizon",
        },
    ],
}
```

## 7. Проверить и собрать PPTX

Сначала выполнить lint:

```python
result = lint_spec(spec)
print(result)
if result["errors"]:
    raise ValueError("\n".join(result["errors"]))
```

Исправить также содержательные предупреждения, если они относятся к текущей
презентации. Затем собрать:

```python
from pathlib import Path

output_dir = Path("/mnt/data") if Path("/mnt/data").exists() else Path.cwd()
path = Path(build_from_spec(spec, output_dir)).resolve()
if (
    not path.is_file()
    or path.suffix.lower() != ".pptx"
    or path.stat().st_size == 0
):
    raise RuntimeError(f"Финальный PPTX не создан: {path}")
print(path, path.stat().st_size)
```

Имя файла задавать латиницей, цифрами, `_` и `-`, без пробелов. Видимое название
может быть на любом языке.

Запустить строгую структурную проверку:

```python
from validate_deck import inspect_deck

qa = inspect_deck(path)
print(qa)
if qa["issues"]:
    raise RuntimeError("\n".join(qa["issues"]))
```

Если доступны office-конвертер и Poppler:

```python
from pathlib import Path
from render_deck import render
from helpers import create_contact_sheet

render_dir = output_dir / "rendered_slides"
slides = render(Path(path), render_dir, dpi=180)
contact_sheet = create_contact_sheet(
    slides,
    output_dir / "contact_sheet.png",
)
print(contact_sheet)
```

Просмотреть каждый слайд отдельно в полном размере и общий contact sheet.
Исправить обрезки, неожиданные переносы, наложения, мелкий текст, слабые crop,
повторяющиеся силуэты, ошибки данных и несогласованные футеры. Структурная
проверка не заменяет визуальный просмотр.

Если системный рендерер отсутствует, экспортировать PPTX в PDF любым доступным
совместимым редактором и проверить растровые страницы. Явно сообщить об
ограничении, если визуальный рендер физически недоступен.

Презентация готова только после выполнения `QUALITY_CHECKLIST.md`.

## 8. Приложить и отдать результат

Сохранить только финальный PPTX и необходимые пользовательские приложения.
Не отдавать рабочий spec, contact sheet и промежуточные файлы, если пользователь
их не просил.

Результат считается отданным только тогда, когда пользователь получает сам
скачиваемый `.pptx`. Текст с именем файла или локальным путём не является
доставкой.

Перед финальным ответом:

1. Убедиться, что `path` указывает на существующий ненулевой `.pptx`.
2. Оставить финальный файл в пользовательском каталоге `/mnt/data`, если он
   доступен.
3. Приложить именно `path` через механизм файловых вложений текущего интерфейса.
4. Если интерфейс использует ссылки `sandbox:`, дать кликабельную ссылку на
   реальный файл, например:

   ```markdown
   [Скачать презентацию](sandbox:/mnt/data/executive_brief.pptx)
   ```

Не придумывать имя или URL в финальном сообщении: ссылка должна указывать на тот
же файл, который вернул `build_from_spec`. Не писать «файл приложен» или
«презентация готова», если скачиваемое вложение фактически не добавлено.

В финальном сообщении:

- первой строкой дать кликабельное вложение или ссылку на финальный `.pptx`;
- кратко назвать результат и количество слайдов;
- указать сюжетную арку и визуальную тему;
- перечислить использованные источники или отметить их отсутствие;
- повторить точное имя созданного `.pptx`.

Минимальный шаблон:

```markdown
[Скачать presentation_name.pptx](sandbox:/mnt/data/presentation_name.pptx)

Готово: презентация на 12 слайдов. Сюжетная арка — проблема → механизм →
решение → следующий шаг. Визуальная тема — ...
```
