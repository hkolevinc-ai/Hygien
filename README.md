# Hygiene.bg → Temu Upload Template Scraper

Този scraper обхожда `https://hygiene.bg/`, взима продуктови данни и попълва sheet `Template` от `TEMU_GENERAL_UPLOAD_TEMPLATE.xlsx`.

## Какво попълва

- `listing_id` — основен SKU/код на продукта или slug, ако няма код
- `seller_sku` — SKU на вариацията; при липса се генерира от `listing_id`
- `item_name` — име на продукта + вариация, когато е приложимо
- `category` — категории от продукта
- `reference_link` — URL на продукта
- `brand` и `trademark` — бранд от сайта
- `Variation_type_1` / `Variation_Variants_1` — обем, аромат, цвят или стандартен вариант
- `Variation_type_2` / `Variation_Variants_2` — втора вариация, ако има
- `item_note` — до 6 кратки bullet точки от описанието
- `item_description` — описание на продукта
- `main_image_url` + `other_image_url` — продуктови изображения в отделни колони
- `map_price` — цена в EUR от сайта
- `Net_content` / `Net_content_unit` — обем/тегло като 500 ml, 1 l, 5 l, ако може да бъде разпознато

## Какво оставя празно по задание

- `quantity`
- package length/width/height/weight
- `item_volume_unit`
- `item_weight_unit`
- `shipping_template`

## Локално пускане

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python scraper.py --mode all_site --limit 20
```

След успешен run ще получиш:

```text
TEMU_HYGIENE_UPLOAD.xlsx
hygiene_raw_export.csv
```

## Пълен run за целия сайт

```bash
python scraper.py --mode all_site --limit 0
```

## Тест с един продукт

```bash
python scraper.py --mode single_product --url "https://hygiene.bg/obezmaslitel-a1/" --limit 1
```

## Тест с категория

```bash
python scraper.py --mode category --url "https://hygiene.bg/cat/pochistvasti-preparati-i-dezinfektanti/" --limit 50
```

## GitHub Actions

1. Качи всички файлове в GitHub repository.
2. Отвори tab **Actions**.
3. Избери workflow **Scrape Hygiene.bg for Temu**.
4. Натисни **Run workflow**.
5. За тест остави `limit: 20`; за целия сайт сложи `limit: 0`.
6. След края свали artifacts: `temu-hygiene-upload`.

## Fix v2

This version skips blog/article pages even when they appear in a sitemap or contain product cards. A page is exported only when it is a real WooCommerce product page (`og:type=product`, `single-product`, or product form/gallery metadata).

Changes:
- product sitemap discovery no longer crawls post/page sitemaps;
- blog pages are ignored during scrape;
- product name is written to `item_name` / Product Name;
- brand is read only from product metadata/product brand blocks, not listing cards;
- full WooCommerce product description is written to `item_description` without the previous 2000-character truncation.
