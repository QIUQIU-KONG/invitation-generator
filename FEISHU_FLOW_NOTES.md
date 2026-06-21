# Feishu Invitation Flow Notes

Daily preview:

```bash
python generate_docs.py --dry-run
```

For non-technical users, double-click:

```text
打开邀请函自动化.vbs
```

If the VBS launcher is blocked by Windows policy, double-click:

```text
打开邀请函自动化.bat
```

The desktop window has buttons for preview, generation, force regeneration, opening Feishu, and opening the output folder.

Template selection:

- Add a `模板类型` field in Feishu or Excel when the operator needs to choose the template manually.
- Supported values are `项目组`, `运营组`, `人事组`, and `供应商`.
- `模板类型` takes priority over the employee department in `knowledge_base.db`.
- If `模板类型` is empty, the script falls back to the employee department from the knowledge base.

Generate after reviewing the preview:

```bash
python generate_docs.py
```

The script will ask for confirmation. Type `YES` to continue.

For unattended runs:

```bash
python generate_docs.py --yes
```

Force regeneration even if the content fingerprint has not changed:

```bash
python generate_docs.py --force --yes
```

Passport rules:

- If Feishu has a passport number, the script treats Feishu as the latest source for that employee.
- If the knowledge base is empty or has an older passport number for the same employee, the script updates `knowledge_base.db` and continues.
- If the passport number already belongs to another employee, the script warns and does not generate that record.
- If both Feishu and the knowledge base are missing the passport number, the script marks the record as `待补护照` and does not generate it.

Only use this when `N/A` is explicitly acceptable:

```bash
python generate_docs.py --allow-missing-passport --yes
```

Incremental generation now uses `record_id` plus a fingerprint of key fields:

- employee id
- name
- department
- position
- gender
- passport
- depart date
- return date

If any key field changes in Feishu, the next run will generate the document again and update `generated.json`.

Older `generated.json` files that only contain `employee_id_depart_date` entries are still recognized, so records generated before the `record_id` upgrade will not be generated again.

When Feishu cannot be read, the default flow stops instead of falling back to the local Excel file. To use Excel intentionally, pass `--excel <file>`.
