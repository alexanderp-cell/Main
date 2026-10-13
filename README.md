# Main — поиск предложений по P/N в Gmail

Скрипт смотрит ваш Gmail (IMAP) и по партийному номеру показывает, какие предложения есть в почте. Результат — своими словами и таблицей под колонки **O–AA** ТУЗ.

## Настройка

1. Файл `.secrets/gmail.env` (уже в `.gitignore`):

```env
GMAIL_ADDRESS=you@domain.com
GMAIL_APP_PASSWORD='xxxx xxxx xxxx xxxx'
```

2. Запуск:

```bash
python3 scripts/pn_quote_search.py 7036410-1904
```

Несколько номеров:

```bash
python3 scripts/pn_quote_search.py 7036410-1904 1590M21P01
```

Столбцом из файла или stdin:

```bash
python3 scripts/pn_quote_search.py --file pns.txt
printf '%s\n' 7036410-1904 1590M21P01 | python3 scripts/pn_quote_search.py --stdin
```

Только TSV (для вставки в Google Sheet / Excel):

```bash
python3 scripts/pn_quote_search.py --tsv-only 7036410-1904
```

## Что заполняется в O–AA

| Col | Поле | Откуда |
|-----|------|--------|
| O | Quote obtained date&time | дата письма |
| P | Initials | вручную |
| Q–R | Root Price / Root | вручную |
| S | Supplier Price per unit, $ | цена из письма |
| T | Supplier | From / имя поставщика |
| U | Transit | вручную |
| V | Cond | NE/OH/SV… |
| W | Lead Time | дни / STOCK |
| X | Qty Avail | qty из письма |
| Y | Cert Info | 8130, EASA Form-1… |
| Z | PPWK | вручную |
| AA | Remarks | короткий фрагмент / вложения |

Парсеры у поставщиков разные — это тренировочный контур: сверяйте вывод с письмом и говорите, что поправить.
