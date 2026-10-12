# Материалы для интерактивного интерьера

В этой папке используются настоящие карты материалов, а не сгенерированные заменители. Все четыре набора имеют CC0 1.0. Они иллюстрируют отделку, но не являются выбранными артикулами поставщиков и не подтверждают стоимость закупки. Светлая древесина показана как визуальное направление: порода дерева не подтверждена.

| Материал | Автор / исходный набор | Лицензия | Карты в HTML |
|---|---|---|---|
| Светлая древесина | Lennart Demes / ambientCG, [Wood048](https://ambientcg.com/view?id=Wood048) | [CC0 1.0](https://creativecommons.org/publicdomain/zero/1.0/) | Color, NormalGL, Roughness |
| Светлая ткань | Lennart Demes / ambientCG, [Fabric036](https://ambientcg.com/view?id=Fabric036) | [CC0 1.0](https://creativecommons.org/publicdomain/zero/1.0/) | Color, NormalGL, Roughness |
| Светлая штукатурка | Lennart Demes / ambientCG, [Plaster001](https://ambientcg.com/view?id=Plaster001) | [CC0 1.0](https://creativecommons.org/publicdomain/zero/1.0/) | Color, NormalGL, Roughness |
| Светлый мрамор | Lennart Demes / ambientCG, [Marble014](https://ambientcg.com/view?id=Marble014) | [CC0 1.0](https://creativecommons.org/publicdomain/zero/1.0/) | Color, NormalGL, Roughness |
| HDRI окружение | HDRI Haven / Poly Haven, [Venice Sunset](https://polyhaven.com/a/venice_sunset) | [CC0](https://polyhaven.com/license) | Radiance HDR, 1024 × 512 |

## Получение и изменения

Официальные домены ambientCG и Poly Haven в текущей сети возвращали ошибку прокси. Материалы ambientCG действительно скачаны из публичного [CC0-зеркала Papyszoo/CC0-Public-Domain-Textures](https://github.com/Papyszoo/CC0-Public-Domain-Textures) с зафиксированной версией `84af0b4614c22c9028143da786fe6f9889bfee70`. Зеркало указывает автора, исходные страницы и CC0. Его README сохранён рядом с картами.

Зеркало уменьшило исходные карты до 512 × 512 и упаковывало их в UASTC KTX2. Здесь старший mip каждого KTX2 декодирован официальным Basis Universal transcoder из Three.js r181; все 12 результатов проверены как читаемые изображения 512 × 512. Цвет сохранён в JPEG (quality 88), нормали — в RGB PNG, шероховатость — в монохромном PNG. Рисунок и оттенки не перекрашивались; новые карты не дорисовывались. JPEG и ранее применённое сжатие KTX2 могут слегка менять пиксели.

HDRI скачано без изменений из [pmndrs/drei-assets](https://github.com/pmndrs/drei-assets/blob/master/hdri/venice_sunset_1k.hdr). README этого зеркала указывает HDRI Haven как источник; копия README сохранена локально. Индивидуальное авторство HDRI здесь не заявляется.

Точные URL загрузки каждой карты, путь, размер, SHA-256, преобразование и лицензия приведены в `materials_manifest.json`. Пути в манифесте отсчитываются от каталога `web/`.

## Применение в Three.js

- Для `color` использовать sRGB, для `normal` и `roughness` — линейные данные. Все материалы диэлектрические, `metalness = 0`.
- Нормали OpenGL: зелёный канал не переворачивать. Начальная сила normal map: древесина 0,25; ткань 0,30; штукатурка 0,12; камень 0,06.
- Рекомендуемый визуальный размер одного повтора: древесина 1 × 2 м; ткань 0,30 × 0,30 м; штукатурка 2 × 2 м; камень 1,2 × 1,2 м. Это настройка визуализации: реальное покрытие исходного снимка не проверено. Масштабировать UV отдельно на каждой поверхности по её размерам.
- Wood048 — рисунок древесины без гарантированного размера доски. Швы напольных планок следует задавать геометрией либо отдельной разметкой, не выдавая число повторов карты за размер паркетной доски.
- HDRI пригодно для отражений и мягкого общего света. Его панорама не изображает реальный вид из окон квартиры; в качестве видимого фона использовать нейтральное окружение.
