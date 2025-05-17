# kaspersky-settings-parser
Скрит для проведения аудита конфигурации АВЗ

Данный скрипт был разработан с цель контроля конфигурации АВЗ
Для скрипта сформированы отдельный Excel файл, в котором хранятся листы,
отражающие часть конфигурации АВЗ.

В листах содержится информация, которая была сочтена полезно для контроля,
в будущем возможно также добавление информации

Скрипт осуществляет контроль конфигурации программ:
 - Kaspersky Endpoint Security для Windows (12.*.0)
 - Kaspersky Security для Windows Server
 - Kaspersky Endpoint Security 12.2 для Linux
 - Kaspersky Endpoint Security для Mac (12.1)
 - Агент администрирования Kaspersky Security Center
 - Kaspersky Security для виртуальных сред 5.2 Легкий агент для Windows

 И соответствующих задач, созданных для текущих программ от производителя Kaspersky.

 Контроль проводится с помощью сбора информации через OpenApi Kaspersky Security Center
