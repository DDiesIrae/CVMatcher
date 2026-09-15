# CV Matcher 0.2 — DeepSeek + OpenAI

Что изменилось:
- выбор AI-провайдера: DeepSeek / OpenAI;
- DeepSeek по умолчанию: `deepseek-flash`;
- кнопка «Проверить подключение»;
- отдельное безопасное хранение ключей через Windows Credential Manager (`keyring`);
- сохранена поддержка Drag & Drop CV;
- старый OpenAI-ключ автоматически подхватывается при переключении на OpenAI.

## DeepSeek
1. Получите API key в кабинете DeepSeek API.
2. Настройки → AI-провайдер → DeepSeek.
3. Вставьте ключ.
4. Оставьте `deepseek-flash`.
5. Нажмите «Проверить подключение» → «Сохранить».

Важно: официальный DeepSeek API тарифицируется отдельно; наличие бесплатного веб-чата DeepSeek не означает бесплатный API.

## GitHub build
После загрузки обновлённых файлов в репозиторий workflow `Build CVMatcher for Windows` запустится автоматически. После зелёной галочки скачайте Artifact `CVMatcher-Windows`.
