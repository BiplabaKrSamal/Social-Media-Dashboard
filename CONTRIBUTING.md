# Contributing to LRZ Social Media Dashboard

## Adding a new platform parser

Each platform follows the same pattern. Here's a minimal example for **TikTok**:

```python
def parse_tiktok(export_path: str | None) -> ParseResult:
    result = ParseResult('tiktok')
    if not export_path or not os.path.exists(export_path):
        result.errors.append("TikTok: export file not found")
        return result

    result.source_files.append(export_path)
    try:
        df = pd.read_csv(export_path)
        df.columns = [c.strip().lower() for c in df.columns]

        # Map TikTok's column names to our schema
        for _, row in df.iterrows():
            d = _parse_date_flexible(row['date'])
            if not d: continue
            mk = _to_month_key(d)
            result.add_month(mk,
                impressions=int(row.get('video views', 0)),
                likes=int(row.get('likes', 0)),
                comments=int(row.get('comments', 0)),
                shares=int(row.get('shares', 0)),
                posts=1
            )
    except Exception as e:
        result.errors.append(f"TikTok parse error: {e}")

    return result
```

Then add it to `run()`:
```python
results['tiktok'] = parse_tiktok(str(data_dir / 'tiktok_export.csv'))
```

And add platform metadata:
```python
PLATFORM_META['tiktok'] = ('TIKTOK', 'FF000000')  # (label, hex color)
```

## PR checklist

- [ ] New parser follows `ParseResult` pattern
- [ ] Sample CSV template in `data/templates/`
- [ ] Export instructions added to `docs/EXPORT_GUIDE.md`
- [ ] Tested with `python lrz_etl.py --dry-run`
- [ ] `lrz_etl_log.json` shows 0 errors for new platform
