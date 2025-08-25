# Быстрый выпуск .exe через GitHub Actions

1. Создай **пустой публичный репозиторий** на GitHub.
2. Зальй сюда все файлы проекта (из архива), обязательно включая:
   - `main.py`, `requirements.txt`, `README_ru.md`, `AvitoPriceAnalyzer.spec`
   - `.github/workflows/windows-build.yml` (этот файл запускает сборку)
3. Создай **тег** и запушь его, чтобы выпустить релиз:
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```
4. Открой вкладку **Actions** → дождись завершения **Build Windows EXE**.
5. Готовый `.exe` будет:
   - во вкладке **Actions** как **artifact**, и
   - во вкладке **Releases** этого репозитория (если пуш был тегом v*).

> Ничего дополнительно настраивать не нужно: используется встроенный `GITHUB_TOKEN`.
