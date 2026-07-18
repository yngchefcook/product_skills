# Product Skills

Скиллы для **продактов**: готовые инструкции агентов под типовые задачи discovery и product work.

Каждый скилл — файл `SKILL.md`. Загружаешь в агента в Yandex AI Studio — и получаешь предсказуемый артефакт (бриф, Lean Canvas, story map и т.д.), а не «кот в мешке» из свободного чата.

Можно использовать **по одному** или собрать в **конвейер** из нескольких агентов.

---

## Как подключить

1. Открой папку скилла → возьми **`SKILL.md`**
2. AI Studio → агент → скиллы → импорт из файла
3. Или вставь содержимое `SKILL.md` как системную инструкцию

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

---

## Что внутри папки скилла

```
skill-slug/
├── SKILL.md      # инструкция агента
└── examples.md   # пример артефакта (опционально)
```