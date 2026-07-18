# Discovery-воркфлоу: как связать скиллы

Мультиагентный конвейер в AI Studio: одна идея продукта → пакет артефактов discovery.

## Схема

```
idea
  └─(1 brief-writing)──────────────▶ brief
         ├─(2 market-research)*───▶ market ──┐
         ├─(3 persona-generation)─▶ personas ┤
         │        └─(7 persona-interview)──▶ interview_report
         ├─(4 lean-canvas)◀──────── brief + market + personas
         ├─(5 user-story-mapping)◀─ brief + personas
         │        └─(6 wireframe-spec)─────▶ wireframes
         └─ все семь артефактов ──▶ презентатор ──▶ discovery_deck

* у ноды 2 включите web search
```

Порядок исполнения: `brief` → (`market` ∥ `personas`) → `interview` → (`lean_canvas`, `story_map`) → `wireframes` → презентатор.

## Правила пайплайна

1. **Не задавать вопросов.** Пробелы во входе → обоснованное допущение с меткой `[assumption]`.
2. **Строгий выход.** Только целевой артефакт в шаблоне скилла (без болтовни до/после).
3. **Единый предмет.** Название продукта и проблема тянутся из `brief` без искажений.
4. **Язык — русский.**

## Реестр артефактов

| id | Продюсер | Суть |
|----|----------|------|
| `idea` | пользователь | 1–2 предложения + опц. ограничения |
| `brief` | brief-writing | задача, Job Story, constraints, критерии, out of scope |
| `market` | market-research | конкуренты, TAM/SAM/SOM, gaps, источники |
| `personas` | persona-generation | 1–3 персоны: цели, боли, возражения |
| `lean_canvas` | lean-canvas | 9 блоков + North Star + гипотезы |
| `story_map` | user-story-mapping | backbone, эпики, релизы |
| `wireframes` | wireframe-spec | карта экранов + unicode-вайрфреймы |
| `interview_report` | persona-interview | вердикт, инсайты, незакрытые возражения |
| `discovery_deck` | ваш презентатор | сводная презентация из семи артефактов |

## Минимальный пайплайн

Если полный конвейер рано:

`brief-writing` → `user-story-mapping` → `wireframe-spec`
