# Product Skills

Скиллы для **продактов**: готовые инструкции агентов под типовые задачи
discovery, product work и создания презентаций.

Каждый скилл начинается с `SKILL.md`. Загружаешь его в агента — и получаешь
предсказуемый артефакт (бриф, Lean Canvas, story map, презентацию и т.д.),
а не «кот в мешке» из свободного чата.

Можно использовать **по одному** или собрать в **конвейер** из нескольких агентов.

---

## Как подключить

1. Открой папку скилла → возьми **`SKILL.md`**
2. Импортируй файл как скилл или вставь его содержимое в системную инструкцию
3. Если внутри есть ресурсы, загрузи их в рабочую среду агента вместе со скиллом

`examples.md` — пример результата, для работы агента не нужен.

---

## Список скиллов

### Основной набор: discovery-конвейер

Из одной идеи продукта — пакет артефактов. Порядок нод:

| # | Скилл | Папка | На выходе |
|---|-------|-------|-----------|
| 1 | Brief Writing | [`brief-writing`](brief-writing/) | бриф |
| 2 | Market Research | [`market-research`](market-research/) | рыночный срез *(нужен web search)* |
| 3 | Persona Generation | [`persona-generation`](persona-generation/) | персоны |
| 4 | Lean Canvas | [`lean-canvas`](lean-canvas/) | Lean Canvas |
| 5 | User Story Mapping | [`user-story-mapping`](user-story-mapping/) | User Story Map |
| 6 | Wireframe Spec | [`wireframe-spec`](wireframe-spec/) | вайрфреймы |
| 7 | Persona Interview | [`persona-interview`](persona-interview/) | отчёт симулированного custdev |

Как связать ноды между собой — [`WORKFLOW.md`](WORKFLOW.md).

```
idea → brief → market ∥ personas → interview
              → lean_canvas + story_map → wireframes
```

**С чего начать:** три ноды  
`brief-writing` → `user-story-mapping` → `wireframe-spec`

### Ещё два скилла

| Скилл | Папка | Для чего |
|-------|-------|----------|
| User Journey Map | [`user-journey-map`](user-journey-map/) | путь пользователя |
| Information Architecture | [`information-architecture`](information-architecture/) | структура продукта / навигация |

### Презентации

| Скилл | Папка | Для чего |
|-------|-------|----------|
| Presentation Code Interpreter | [`presentation-code-interpreter`](presentation-code-interpreter/) | проектирование, сборка и проверка `.pptx` |

Это расширенный скилл с готовым Python-конструктором, 14 типами слайдов,
4 визуальными темами, автоматической проверкой и 10 нейтральными изображениями.
Он умеет использовать как локальные ресурсы, так и новые визуалы через
`image_gen.imagegen` или совместимый `generate-image`.

Как запустить:

1. Вставь [`SKILL.md`](presentation-code-interpreter/SKILL.md) в системную
   инструкцию агента.
2. Загрузи содержимое папки
   [`code_interpreter`](presentation-code-interpreter/code_interpreter/) в
   Code Interpreter, сохранив структуру файлов.
3. Передай задачу, аудиторию, исходные материалы и желаемый формат презентации.

---

## Что внутри папки скилла

Обычный текстовый скилл:

```
skill-slug/
├── SKILL.md      # инструкция агента
└── examples.md   # пример артефакта (опционально)
```

Скилл с исполняемыми ресурсами:

```
skill-slug/
├── SKILL.md
└── code_interpreter/
    ├── builder.py
    ├── helpers.py
    ├── image_filenames.py
    ├── images/
    ├── examples/
    └── документация и конфигурация
```
