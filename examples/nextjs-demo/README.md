# gridpath in Next.js

The smallest Next.js app (App Router) with a `gridpath` spreadsheet and the
server half of an agent loop:

- `app/Sheet.tsx`: the spreadsheet, a client component (its engine runs in
  a Web Worker in the browser).
- `app/page.tsx`: a server component that renders it.
- `app/api/agent/route.ts`: one Claude call per request with the tools and
  instructions from `gridpath/agent`; the key stays on the server
  (`ANTHROPIC_API_KEY`). The browser half of the loop is the same as in
  [`../embed-demo/src/agentLoop.ts`](../embed-demo/src/agentLoop.ts).

```sh
(cd ../../packages/gridpath && npm install && npm pack)   # the package, from source
npm install
npm run dev            # or: npm run build && npm start (add --webpack to build with webpack)
```

No Next.js configuration is needed with Turbopack or webpack.
