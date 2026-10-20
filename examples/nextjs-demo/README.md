# gridpath in Next.js

A small but complete Next.js app (App Router): each signed-in user has a
list of workbooks saved to the app's own API, and a spreadsheet to edit
them.

- **Loading by document ID.** Each workbook has a URL, `/books/<id>`. The
  spreadsheet gets `documentId={id}` and an `httpStore`, and opens that
  workbook from the API. A new id opens a blank workbook.
- **Saving.** Saves go to the API automatically as the user works. They
  are sent with the user's token, and the status line says "Saved to your
  account".
  - If the API can't be reached, the edits are kept in this browser
    ("Saved in this browser only") and the save is retried. If the user
    closes the tab first, the edits reopen from the browser next time.
- **Uploads and new workbooks.** An uploaded `.xlsx` is opened by the
  spreadsheet, which checks it, and then saved to the API.
- **Switching workbooks.** Moving to another workbook remounts the
  spreadsheet: it saves the one being left, stops its engine, and opens the
  next.
- **Loading indicator and errors.** The page shows the loading stages
  (`onProgress`). Problems arrive at `onError` as a `GridPathError` with a
  `code`, and the app adds its own way out per code. For example, a CSV
  upload shows the spreadsheet's explanation plus a "Back to your
  workbooks" link.

| File | What it is |
| --- | --- |
| `app/Workbooks.tsx` | The client component: the workbook list, New, Upload, the spreadsheet, the loading overlay, and the per-code error actions. |
| `app/page.tsx`, `app/books/[id]/page.tsx` | Server components that render it. |
| `app/api/books/[[...path]]/route.ts` | The REST API `httpStore()` calls: `PUT`, `GET` and `DELETE` per workbook, the list, and version history. |
| `lib/books.ts` | Storage on disk (`.data/`, a folder per user). Replace it with your database or object storage. |
| `lib/auth.ts`, `app/session.ts` | Stand-ins for your auth: a bearer token checked by the route. With cookie sessions, `httpStore` needs no `headers`. |
| `app/api/agent/route.ts` | The server half of an AI agent loop (see the gridpath README). The key stays on the server (`ANTHROPIC_API_KEY`). |

```sh
(cd ../../packages/gridpath && npm install && npm pack)   # the package, from source
npm install
npm run dev            # or: npm run build && npm start
```

No Next.js configuration is needed, with either Turbopack or webpack.
