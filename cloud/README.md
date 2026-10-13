# Online PDF processing

New decisions are processed in GitHub Actions, so nothing needs to run on anyone's computer.

## How it works

1. An admin signs in to the website and uses **Add PDFs**. Files go to the private `incoming` bucket in Supabase Storage.
2. Every morning (about 5 AM Eastern), or whenever someone runs the workflow by hand, `cloud_process.py`:
   - summarizes each new PDF with Claude,
   - saves it in the `case-pdfs` bucket (in a year folder) and in the `documents` table,
   - rebuilds the Word summary (a clickable link to every PDF) and the PDF summary.
3. Members see the new decisions in search, and the Word and PDF download buttons show the latest summary.

## One-time setup (the new owner does this under their own accounts)

In the GitHub repo: **Settings > Secrets and variables > Actions > New repository secret**. Add three secrets:

| Name | Value |
| --- | --- |
| `ANTHROPIC_API_KEY` | The owner's own Anthropic API key |
| `SUPABASE_URL` | `https://dhsjjsauzixhwxsctjnn.supabase.co` (stays the same after a project transfer) |
| `SUPABASE_SERVICE_KEY` | The `service_role` secret key from Supabase > Project Settings > API |

The service key bypasses all access rules. Keep it only in GitHub Secrets.

## Running it by hand

Repo > **Actions** > **Process new PDFs** > **Run workflow**. Tick "Rebuild" to regenerate the Word and PDF summaries even when nothing new is waiting. This also works from the GitHub phone app.

## Making someone an admin

Run this in the Supabase SQL editor (replace the email):

```sql
update auth.users
set raw_app_meta_data = coalesce(raw_app_meta_data, '{}'::jsonb) || '{"role":"admin"}'::jsonb
where email = 'person@example.org';
```

They sign out and back in to see the **Add PDFs** button.

## If something goes wrong

- A run that shows a red X lists which file failed and why on the run's summary page. Failed files stay in `incoming` and are retried at the next run.
- PDFs already in the archive are moved to `incoming/_skipped/` instead of being summarized again.
- `case_summarizer_supabase.py` here is the shared code. It is the same file as the laptop version, and `cloud_process.py` reuses its functions.
