# Восстановление нового экспорта Claude из manifest и ZIP-файлов

Этот документ — самодостаточная инструкция для человека или LLM-агента, который впервые видит этот каталог. Он описывает полный процесс: распаковку нового экспорта Claude, восстановление структуры диалогов и созданных в них файлов, попытку повторной генерации PPTX/DOCX/XLSX, обновление HTML-интерфейса и итоговую проверку.

Поддерживаются **Linux/WSL** и **Windows** (с нативным Python, Node.js и инструментами, установленными через winget).

Интерфейс поддерживает несколько языков: русский (по умолчанию), китайский и английский. Перевод инструкций на китайский: `RESTORE_NEW_CLAUDE_EXPORT_zh.md`.

## 1. Что находится в этом каталоге

Рабочие скрипты:

- `restore_export.py` — основной восстановитель и генератор HTML-интерфейса;
- `regenerate_python_office.py` — безопасная попытка запуска восстановленных Python-генераторов Office-файлов;
- `regenerate_node_office.py` — попытка запуска восстановленных Node.js-генераторов PPTX/DOCX/XLSX;
- `regenerate_recipe_docx.py` — воспроизведение многошаговых DOCX-рецептов, сохранённых в tool calls (TeX, рисунки, Pandoc и `python-docx`);
- `regenerate_known_office.py` — воспроизведение явно проверенных рецептов и точных операций переименования для известных DOCX/XLSX этого экспорта;
- `artifact_timestamps.py` — восстановление `mtime` перегенерированных артефактов по журналам выдачи файлов и tool calls;
- `check_restoration.py` — проверка локальных ссылок и целостности полученных Office-файлов.

Вспомогательные файлы:

- `winenv.py` — кроссплатформенный модуль для автоматического обнаружения инструментов, установленных через winget на Windows, и корректной подстановки путей в генерируемый исходный код;
- `install_deps_windows.ps1` — PowerShell-скрипт для автоматической установки всех системных и языковых зависимостей на Windows через winget, pip и npm;
- `restore_all.bat` — один командный файл для запуска всего процесса на Windows от первого прохода до проверки.

Главный результат — каталог, переданный через `--output`. В примерах ниже он называется `restored_export_new`.

## 2. Что требуется на входе

Поместите в один исходный каталог:

- новый файл `manifest-....json`;
- все относящиеся к нему `*.zip`;
- ничего не переименовывайте внутри ZIP-архивов.

Пример:

```text
new_export/
├── manifest-xxxxxxxx.json
├── part-01.zip
├── part-02.zip
└── part-03.zip
```

Важные детали:

1. `restore_export.py` ищет `*.zip` только непосредственно в каталоге `--source`, не рекурсивно.
2. Имя manifest-файла может быть любым. Текущая версия восстановителя не использует manifest как источник содержимого: manifest нужен как опись комплекта, а данные извлекаются из всех ZIP-файлов каталога.
3. После распаковки среди данных должен существовать `conversations.json`. Обычно он приходит из одного из ZIP-архивов.
4. Не смешивайте ZIP-файлы от разных экспортов в одном `--source`: восстановитель распакует их все в общий `raw/`.
5. Для каждого нового комплекта используйте новый пустой выходной каталог. Это исключает смешивание старых и новых диалогов.

## 3. Требования к среде

### 3.1. Минимальные требования (все платформы)

- Python 3.10 или новее;
- пакет `markdown-it-py` для HTML-рендеринга Markdown;
- достаточно свободного места: результат может быть значительно больше суммарного размера ZIP.

### 3.2. Для перегенерации Office-файлов (опционально)

Дополнительные пакеты и инструменты нужны только если вы хотите повторно сгенерировать PPTX/DOCX/XLSX из восстановленных исходников и команд. Чаты, текст и метаданные восстанавливаются без них.

**Python-библиотеки** (все платформы):

```bash
python -m pip install markdown-it-py openpyxl python-docx python-pptx matplotlib pandas numpy pillow
```

**Node.js** (для PPTX/DOCX-генераторов на `pptxgenjs` и `docx`):

- `node` 18+ и `npm`;
- пакеты: `pptxgenjs@3.12.0`, `docx@9.5.1`, `react`, `react-dom`, `react-icons`, `sharp`.

**Системные инструменты** (для DOCX-рецептов с TeX/Pandoc):

- `pandoc` — преобразование Markdown/LaTeX в DOCX;
- `pdflatex` (TeX Live или MiKTeX) — компиляция TikZ/PGFPlots-рисунков;
- `pdftoppm` (Poppler) — преобразование PDF-рисунков в PNG;
- `libreoffice`/`soffice` — дополнительная визуальная проверка DOCX.

### 3.3. Установка зависимостей на Windows

#### Автоматическая установка (рекомендуется)

Запустите PowerShell-скрипт из каталога `recover_chats`:

```powershell
powershell -ExecutionPolicy Bypass -File recover_chats\install_deps_windows.ps1
```

Скрипт в silent-режиме (без интерактивного ввода) установит через `winget`:

| Инструмент | winget ID | Назначение |
|---|---|---|
| Pandoc 3.14 | `JohnMacFarlane.Pandoc` | Markdown/LaTeX → DOCX |
| Poppler 25.07 | `oschwartz10612.Poppler` | PDF → PNG (`pdftoppm`) |
| MiKTeX 25.12 | `MiKTeX.MiKTeX` | `pdflatex`, TikZ, PGFPlots |
| LibreOffice 26.8 | `TheDocumentFoundation.LibreOffice` | Проверка DOCX → PDF |

И через `pip` — все Python-библиотеки.

После первого прохода `restore_export.py` установите npm-пакеты:

```cmd
npm install --prefix restored_export_new\runtime pptxgenjs@3.12.0 docx@9.5.1 react react-dom react-icons sharp --no-audit --no-fund
```

#### Ручная установка по одной команде

```cmd
winget install --id JohnMacFarlane.Pandoc --accept-package-agreements --accept-source-agreements
winget install --id oschwartz10612.Poppler --accept-package-agreements --accept-source-agreements
winget install --id MiKTeX.MiKTeX --accept-package-agreements --accept-source-agreements
winget install --id TheDocumentFoundation.LibreOffice --accept-package-agreements --accept-source-agreements
```

Флаги `--accept-package-agreements --accept-source-agreements` делают установку полностью неинтерактивной.

#### Как скрипты находят инструменты на Windows

Инструменты, установленные через winget, помещаются в per-user каталоги (`%LOCALAPPDATA%`), которые не всегда добавляются в `PATH` в неинтерактивных сессиях. Модуль `winenv.py` автоматически:

1. Находит Pandoc в `%LOCALAPPDATA%\Pandoc\`;
2. Находит MiKTeX в `%LOCALAPPDATA%\Programs\MiKTeX\miktex\bin\x64\`;
3. Находит Poppler в winget-каталоге пакетов по шаблону `oschwartz10612.Poppler_*`;
4. Находит LibreOffice в `C:\Program Files\LibreOffice\program\`;
5. Добавляет найденные каталоги в `PATH` текущего процесса Python.

Все скрипты (`regenerate_*.py`) импортируют `winenv` и вызывают `ensure_tools_on_path()` при запуске — вручную настраивать `PATH` не нужно.

### 3.4. Установка зависимостей на Linux/WSL

```bash
sudo apt-get update
sudo DEBIAN_FRONTEND=noninteractive apt-get install -y \
  pandoc \
  texlive-latex-extra \
  texlive-pictures \
  texlive-fonts-recommended \
  texlive-extra-utils \
  poppler-utils \
  libreoffice-core-nogui \
  libreoffice-writer-nogui \
  default-jre-headless
```

Python:

```bash
python3 -m pip install markdown-it-py openpyxl python-docx python-pptx matplotlib pandas numpy pillow cairosvg
```

Node.js:

```bash
npm install --prefix restored_export_new/runtime pptxgenjs@3.12.0 docx@9.5.1 react react-dom react-icons sharp --no-audit --no-fund
```

Назначение пакетов:

- `pandoc` преобразует подготовленный LaTeX в DOCX;
- `texlive-latex-extra`, `texlive-pictures` и `texlive-fonts-recommended` дают `pdflatex`, TikZ, PGFPlots и шрифтовые пакеты;
- `texlive-extra-utils` содержит `pdfcrop`;
- `poppler-utils` предоставляет `pdftoppm`, `pdfinfo` и `pdftotext`;
- `libreoffice-*-nogui` позволяет проверять DOCX конвертацией в PDF без X11;
- `default-jre-headless` нужен части компонентов LibreOffice.

На сервере и в WSL устанавливайте именно `libreoffice-core-nogui` и `libreoffice-writer-nogui`. Обычный GUI-пакет `libreoffice-writer` может пытаться открыть X11 даже с параметром `--headless`.

### 3.5. Особенности Windows, учтённые в скриптах

1. **Escape-последовательности в путях.** При подстановке Windows-путей с обратными слешами в исходный код генераторов (`\a` в `\anthropic`, `\f` в `\flu`, `\2` в `\2026`) Python и JavaScript интерпретируют их как escape-последовательности. Все скрипты используют `to_posix_path()` из `winenv.py`, который возвращает пути с прямыми слешами.

2. **Блокировка файлов pdflatex.** На Windows `pdflatex` может оставлять `.log`-файлы открытыми, что препятствует очистке временного каталога (`PermissionError: [WinError 32]`). Функция `kill_orphaned_tex_processes()` вызывается в блоке `finally` после LaTeX-компиляций.

3. **`python3` vs `python`.** На Windows исполняемый файл Python обычно называется `python`, а не `python3`. Все скрипты используют `sys.executable` вместо жёстко заданного `python3`.

4. **`cairosvg` без нативной Cairo.** На Windows `cairosvg` устанавливается через pip, но требует нативную библиотеку `libcairo-2.dll`, которая отсутствует. DOCX-рецепты со SVG-диаграммами обрабатываются через альтернативный путь в `regenerate_known_office.py`; `cairosvg` используется только при его успешном импорте.

## 4. Безопасность

ZIP-файлы и восстановленные генераторы следует считать недоверенными.

Основной восстановитель:

- проверяет CRC каждого ZIP;
- отклоняет абсолютные пути и компоненты `..` внутри ZIP;
- не запускает извлечённые команды при обычном восстановлении;
- сохраняет команды и исходники как материалы аудита.

Скрипты перегенерации действительно запускают часть восстановленного Python/JavaScript-кода. Они отбрасывают известные опасные конструкции и перенаправляют типовые выходные пути в `regenerated/`, но это эвристическая, а не формальная песочница.

`regenerate_recipe_docx.py` также выполняет восстановленные промежуточные Python-скрипты, но копирует их и исходники в одноразовый временный каталог. В каталог диалога копируется только итоговый проверенный DOCX.

`regenerate_known_office.py` содержит явный перечень проверенных рецептов этого набора. Он не является универсальным распознавателем: для другого экспорта перечень следует расширять только после аудита соответствующих tool calls и входных файлов.

Поэтому перегенерацию следует выполнять:

- в контейнере, VM или отдельном непривилегированном окружении;
- без секретов и рабочих токенов в переменных среды;
- без доступа к важным каталогам;
- желательно без сети;
- только после просмотра списков кандидатов или с пониманием этого риска.

Если запуск недоверенного кода недопустим, пропустите разделы 6 и 7. Чаты, исходники, команды, вложенный текст и метаданные всё равно будут восстановлены.

## 5. Первый проход: распаковка и восстановление

Все команды далее выполняются из каталога, где лежат актуальные Python-скрипты этого проекта.

**Linux/WSL:**

```bash
python3 restore_export.py \
  --source /absolute/path/to/new_export \
  --output /absolute/path/to/restored_export_new
```

**Windows:**

```cmd
python recover_chats\restore_export.py --source . --output restored_export_new
```

Что делает этот проход:

1. Находит все `*.zip` непосредственно в `--source`.
2. Проверяет CRC и безопасность путей.
3. Распаковывает данные в `<output>/raw/`.
4. Читает `<output>/raw/conversations.json`.
5. Для каждого диалога создаёт отдельный каталог.
6. Восстанавливает состояния файлов из операций `Write`, `Edit`, `create_file`, `str_replace`, heredoc и поддерживаемых shell-команд.
7. Сохраняет команды, inline-скрипты и извлечённый текст вложений.
8. Генерирует `chat.md`, `chat.html`, JSON-манифесты и общий `INDEX.html`.
9. Восстанавливает проектные документы и каталог Claude Artifacts, если они есть в экспорте.

Нормальный успешный код возврата — `0`. Код `2` означает, что после распаковки не найден `raw/conversations.json`.

### Что проверить после первого прохода

Откройте:

```text
<output>/archive_report.json
<output>/SUMMARY.json
<output>/INDEX.html
```

В `archive_report.json` у каждого архива поле `error` должно быть `null`. Если архив повреждён, не продолжайте молча: замените его правильной копией и начните с нового выходного каталога.

## 6. Попытка перегенерации Python Office-файлов

**Linux/WSL:**

```bash
python3 regenerate_python_office.py /absolute/path/to/restored_export_new/conversations
```

**Windows:**

```cmd
python recover_chats\regenerate_python_office.py restored_export_new\conversations
```

Скрипт ищет восстановленные `*.py`, которые:

- используют `openpyxl`, `python-docx` или `python-pptx`;
- объявляют выходной PPTX/DOCX/XLSX;
- не содержат распознанных сетевых, shell- или разрушительных вызовов.

Созданные файлы помещаются в:

```text
<conversation>/regenerated/
```

Отчёт записывается в:

```text
<output>/python_regeneration_report.json
```

Ненулевой код возврата здесь не обязательно означает провал всего восстановления. Он означает, что хотя бы один подходящий генератор завершился с ошибкой. Изучите для каждой записи поля `status`, `exit_code` и `log`. Типичные причины:

- исходный генератор рассчитывал на отсутствующее изображение или таблицу;
- в экспорте нет базового бинарного файла;
- генератор имел ошибку уже в исходном диалоге;
- отсутствует дополнительная Python-библиотека.

Не устанавливайте неизвестные зависимости автоматически без анализа генератора.

## 7. Попытка перегенерации Node.js Office-файлов

Сначала установите зависимости из раздела 3, затем выполните:

**Linux/WSL:**

```bash
python3 regenerate_node_office.py \
  restored_export_new/conversations \
  restored_export_new/runtime/node_modules
```

**Windows:**

```cmd
python recover_chats\regenerate_node_office.py restored_export_new\conversations restored_export_new\runtime\node_modules
```

Скрипт ищет восстановленные JavaScript-генераторы на основе `pptxgenjs` и `docx`, перенаправляет известные выходные пути в `regenerated/` и пропускает распознанные опасные или дублирующиеся генераторы.

Отчёт записывается в:

```text
<output>/node_regeneration_report.json
```

Как и для Python, отдельные `failed`, `timeout` или `skipped-*` допустимы и должны быть объяснены по журналу. Не удаляйте успешно созданные файлы из-за ошибок других генераторов.

## 8. DOCX-рецепты и обязательное обновление HTML

### 8.1. Перегенерация DOCX по сохранённым tool-call-рецептам

После установки системных зависимостей выполните:

**Linux/WSL:**

```bash
python3 regenerate_recipe_docx.py restored_export_new/conversations
```

**Windows:**

```cmd
python recover_chats\regenerate_recipe_docx.py restored_export_new\conversations
```

Supported recipes use recovered local inputs.

1. копирование `paper.tex`, файлов рисунков и восстановленных `prep.py`, `prep2.py`, `post.py`;
2. применение поздних исправлений, записанных последующими `sed`/Python tool calls;
3. компиляция TikZ/PGFPlots-рисунков через `pdflatex`;
4. преобразование рисунков и формул в PNG;
5. сборка промежуточного DOCX через Pandoc;
6. оформление через `python-docx`;
7. проверка ZIP/CRC и повторное чтение через `python-docx`;
8. дополнительная попытка конвертации LibreOffice в PDF.

Итоговые документы помещаются в `<conversation>/regenerated/`, а отчёт сохраняется в:

```text
<output>/recipe_docx_regeneration_report.json
```

Успешная запись должна иметь `status: "ok"`. Поле `validation` показывает CRC, число абзацев, таблиц, встроенных изображений и результат дополнительной проверки LibreOffice.

### 8.2. Проверенные рецепты DOCX/XLSX конкретного экспорта

Для этого репозитория дополнительно выполните:

**Linux/WSL:**

```bash
python3 regenerate_known_office.py \
  restored_export_new/conversations \
  restored_export_new/runtime/node_modules
```

**Windows:**

```cmd
python recover_chats\regenerate_known_office.py restored_export_new\conversations restored_export_new\runtime\node_modules
```

Скрипт воспроизводит только явно сопоставленные рецепты: Pandoc из восстановленного Markdown, вызовы восстановленного `md2docx.js`, полные Node.js-рецепты с диаграммами и точные `cp`/rename-операции из tool calls. Каждый результат проверяется как ZIP/CRC и повторно открывается через `python-docx` или `openpyxl`. Отчёт: `<output>/known_office_regeneration_report.json`.

Если нужное имя присутствует только в `presented_files.json`, но отсутствует финальный исходник, шаблон, изображение или набор данных, скрипт не создаёт фиктивную замену. Такой файл остаётся в списке невосстановленных с объяснимой причиной.

### 8.3. Обновление HTML после всех перегенераций

После любой перегенерации повторите основной запуск с теми же `--source` и `--output`:

**Linux/WSL:**

```bash
python3 restore_export.py \
  --source /absolute/path/to/new_export \
  --output /absolute/path/to/restored_export_new
```

**Windows:**

```cmd
python recover_chats\restore_export.py --source . --output restored_export_new
```

Этот проход нужен, чтобы:

- обнаружить уже созданные файлы в `regenerated/`;
- связать их с упоминаниями в соответствующих сообщениях;
- добавить их в правую панель артефактов;
- присвоить им даты из логов;
- обновить общий индекс и HTML каждого чата.

Основной восстановитель не удаляет `regenerated/`, поэтому успешно созданные Office-файлы сохраняются.

Если перегенерация была полностью пропущена, второй проход не обязателен.

### 8.4. Восстановление времени файлов

Все скрипты перегенерации после успешного создания артефакта вызывают `artifact_timestamps.py` и восстанавливают время изменения (`mtime`) результирующего файла из журнала диалога:

1. Для финального файла выбирается последнее событие `present_files`/`SendUserFile` с точно совпадающим именем.
2. Если промежуточный файл не выдавался пользователю, используется последнее сообщение/tool call, содержащее его полное имя.
3. Источник, timestamp и UUID сообщения записываются в поле `restored_mtime` или `restored_mtimes` соответствующего отчёта перегенерации.

Последнее событие используется намеренно: один и тот же путь мог выдаваться несколько раз, а перегенерированный файл соответствует его финальной версии. Время доступа (`atime`) сохраняется. Переносимое изменение времени создания/birth time невозможно, поэтому на Linux и Windows гарантированно восстанавливается именно `mtime`.

Для уже созданного каталога результат можно обновить без повторного запуска генераторов:

```bash
python3 artifact_timestamps.py restored_export_new/conversations
```

На Windows:

```cmd
python recover_chats\artifact_timestamps.py restored_export_new\conversations
```

## 9. Итоговая проверка

**Linux/WSL:**

```bash
python3 check_restoration.py /absolute/path/to/restored_export_new
```

**Windows:**

```cmd
python recover_chats\check_restoration.py restored_export_new
```

Критерии успеха:

- `broken_links` равно `0`;
- `corrupt_office_files` равно `0`;
- `html_files` соответствует индексу, индексу Artifacts и числу разговоров;
- `office_files` соответствует реально перегенерированным файлам.

Полный отчёт сохраняется в:

```text
<output>/validation_report.json
```

Проверка Office-файлов проверяет ZIP-структуру и CRC для PPTX/DOCX/XLSX. Она не гарантирует смысловую правильность содержимого слайдов, документов или формул.

## 10. Как пользоваться результатом

Откройте в браузере:

```text
<output>/INDEX.html
```

Интерфейс содержит:

- слева — список диалогов по дате последней активности;
- переключатель светлой и тёмной темы с сохранением выбора;
- поиск и сортировку диалогов;
- в центре — выбранный HTML-чат с Markdown-рендерингом;
- справа — артефакты с датами, версиями и сортировкой;
- предпросмотр PDF, изображений, HTML и текстовых файлов;
- ссылки на PPTX/DOCX/XLSX;
- восстановленные файлы с датой создания, датой последнего изменения и числом ревизий;
- сортировку восстановленных файлов по дате, имени и числу ревизий.

Если браузер ограничивает вложенные локальные страницы при открытии через `file://`, запустите простой локальный HTTP-сервер:

```bash
python3 -m http.server 8000 \
  --bind 127.0.0.1 \
  --directory /absolute/path/to/restored_export_new
```

После этого откройте:

```text
http://127.0.0.1:8000/INDEX.html
```

HTTP-сервер следует привязывать только к loopback-интерфейсу и не публиковать в сеть, если экспорт содержит конфиденциальные данные.

## 11. Многоязычная локализация (i18n)

Интерфейс поддерживает несколько языков. Текущие языки: русский (`ru`, по умолчанию), китайский (`zh`), английский (`en`).

### Как это работает

- `restore_export.py` генерирует отдельный HTML-файл для каждого языка: `chat.html` (по умолчанию), `chat_en.html`, `chat_zh.html` и `INDEX.html`, `INDEX_en.html`, `INDEX_zh.html`.
- В сайдбаре `INDEX.html` есть переключатель языка (выпадающий список).
- При смене языка браузер сохраняет выбранное значение в cookie `restore_lang` (срок 1 год) и перенаправляет на соответствующую HTML-страницу.
- JavaScript-сортировка диалогов и артефактов использует локаль выбранного языка (`localeCompare` с правильным BCP-47 кодом).

### Как добавить новый язык

1. Создайте файл `locales/<lang>.json` по образцу существующих (`ru.json`, `zh.json`, `en.json`).
2. Заполните все ключи. Структура — вложенные объекты: `index.*`, `chat.*`, `artifacts_index.*`, плюс `lang_code`, `lang_name`, `html_lang`, `sort_locale`.
3. Повторно запустите `restore_export.py` — новый язык автоматически подхватится из каталога `locales/`.

Файлы локализации:

```text
recover_chats/locales/
├── ru.json    (русский, по умолчанию)
├── zh.json    (中文)
└── en.json    (English)
```

### Перевод инструкций

- `RESTORE_NEW_CLAUDE_EXPORT.md` — русская версия (основная).
- `RESTORE_NEW_CLAUDE_EXPORT_zh.md` — китайская версия.

## 12. Структура результата

```text
restored_export_new/
├── INDEX.html          (русский, по умолчанию)
├── INDEX_en.html       (английский)
├── INDEX_zh.html       (китайский)
├── INDEX.md
├── SUMMARY.json
├── archive_report.json
├── restoration_report.json
├── validation_report.json
├── python_regeneration_report.json
├── node_regeneration_report.json
├── recipe_docx_regeneration_report.json
├── known_office_regeneration_report.json
├── raw/
├── project_documents/
└── conversations/
    └── <date>_<title>_<uuid>/
        ├── chat.html
        ├── chat.md
        ├── conversation.json
        ├── README.md
        ├── recovered_files.json
        ├── presented_files.json
        ├── files/
        ├── commands/
        ├── inline_scripts/
        ├── attachments/
        └── regenerated/
```

Не каждый каталог обязательно присутствует в каждом разговоре.

## 13. Ограничения формата экспорта Claude

Следует различать три случая:

1. **Файл присутствует в ZIP.** Его можно извлечь точно.
2. **В логах сохранился исходник или генератор.** Файл можно попытаться воспроизвести, но он не обязательно будет побайтно идентичен оригиналу.
3. **Сохранились только путь, UUID и MIME-тип.** Бинарный файл восстановить невозможно без внешней копии; интерфейс покажет, что доступны только метаданные.

Экспорт часто не содержит байты файлов, которые Claude показал пользователю через `present_files`. Поэтому наличие имени `report.pptx` в чате не доказывает наличие самого PPTX в архиве.

Извлечённый текст вложения сохраняется с окончанием `.extracted.txt`, даже если исходное вложение было PDF или DOCX. Это сделано намеренно: текстовая расшифровка не должна маскироваться под отсутствующий бинарный оригинал.

## 14. Полная последовательность команд

### 13.1. Windows (одной командой)

```cmd
recover_chats\restore_all.bat . restored_export_new
```

Или вручную по шагам:

```cmd
powershell -ExecutionPolicy Bypass -File recover_chats\install_deps_windows.ps1

python recover_chats\restore_export.py --source . --output restored_export_new

python recover_chats\regenerate_python_office.py restored_export_new\conversations

npm install --prefix restored_export_new\runtime pptxgenjs@3.12.0 docx@9.5.1 react react-dom react-icons sharp --no-audit --no-fund

python recover_chats\regenerate_node_office.py restored_export_new\conversations restored_export_new\runtime\node_modules

python recover_chats\regenerate_recipe_docx.py restored_export_new\conversations

python recover_chats\regenerate_known_office.py restored_export_new\conversations restored_export_new\runtime\node_modules

python recover_chats\restore_export.py --source . --output restored_export_new

python recover_chats\check_restoration.py restored_export_new
```

### 13.2. Linux/WSL

```bash
python3 restore_export.py \
  --source /data/new_export \
  --output /data/restored_export_new

python3 regenerate_python_office.py \
  /data/restored_export_new/conversations

npm install \
  --prefix /data/restored_export_new/runtime \
  pptxgenjs@3.12.0 docx@9.5.1 react react-dom react-icons sharp \
  --no-audit --no-fund

python3 regenerate_node_office.py \
  /data/restored_export_new/conversations \
  /data/restored_export_new/runtime/node_modules

python3 regenerate_recipe_docx.py \
  /data/restored_export_new/conversations

python3 regenerate_known_office.py \
  /data/restored_export_new/conversations \
  /data/restored_export_new/runtime/node_modules

python3 restore_export.py \
  --source /data/new_export \
  --output /data/restored_export_new

python3 check_restoration.py \
  /data/restored_export_new
```

Если один из скриптов перегенерации вернул ненулевой код, прочитайте соответствующий JSON-отчёт, зафиксируйте причины ошибок и продолжайте с сохранившимися успешными результатами. Не объявляйте весь экспорт невосстановимым только из-за отдельных генераторов.

## 15. Инструкция для автономного ЛЛМ-агента

Если эту работу выполняет другая LLM-модель, она должна:

1. Определить абсолютные пути исходного и нового выходного каталогов.
2. Убедиться, что в исходном каталоге находятся все ZIP одного manifest-комплекта.
3. Не смешивать новый результат с ранее восстановленным экспортом.
4. Проверить доступность Python и `markdown_it`.
5. Выполнить первый проход `restore_export.py` и изучить `archive_report.json`.
6. Не запускать недоверенные генераторы без допустимой изоляции.
7. При разрешённой перегенерации запустить Python-, Node-, общий DOCX-рецепт и проверенные Office-рецепты по отдельности и сохранить их отчёты.
8. Выполнить второй проход `restore_export.py` после перегенерации.
9. Запустить `check_restoration.py` и не завершать задачу, пока не объяснены битые ссылки или повреждённые Office-файлы.
10. Сообщить пользователю абсолютный путь к `INDEX.html`, число диалогов, файлов, артефактов, успешных перегенераций и оставшиеся ограничения.
11. Не удалять исходные ZIP, manifest, отчёты, восстановленные генераторы или неудачные журналы без отдельного прямого указания пользователя.
