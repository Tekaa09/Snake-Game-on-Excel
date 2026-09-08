# 🐍 Snake in Excel (VBA)

A classic Snake game built entirely inside Microsoft Excel using VBA — no
external engine, just worksheet cells, colors, and macros.

I built this project while studying for the Microsoft Office Specialist
(MOS) certification in 10th grade, as a way to explore just how much you
could actually do inside Excel beyond spreadsheets and formulas.

## How it works

The game grid is a block of worksheet cells (`B2:Q34`). The snake is drawn
by coloring cells: on every "tick" of a VBA timer, the snake's body shifts
forward by one cell, the game checks for collisions, and the grid is
repainted. Apples (yellow cells) and obstacles (blue cells) spawn at
random empty positions on the grid.

| Module  | Role                                                        |
|---------|-------------------------------------------------------------|
| `Main`  | Game state, movement, collision detection, apples, scoring  |
| `Keys`  | Arrow-key bindings that control the snake's direction       |
| `Enemy` | A second, independently-moving snake                        |

## Controls

- Arrow keys — change direction
- Run the `StartGame` macro (or its button on the sheet) to begin

## Getting started

1. Download `Snake11.xlsm`.
2. Open it in Excel and enable macros when prompted (*Enable Content*).
3. Run the `StartGame` macro, or click the start button on the sheet.

> The `source/` folder contains the raw VBA module code (`Main.bas`,
> `Keys.bas`, `Enemy.bas`) exported from the workbook, for anyone who
> wants to read the code without opening Excel.

## Project structure

```
.
├── Snake11.xlsm       # The playable workbook
├── source/             # VBA modules exported as plain text, for reference
│   ├── Main.bas
│   ├── Keys.bas
│   └── Enemy.bas
└── README.md
```

## License

Released under the [MIT License](LICENSE).
