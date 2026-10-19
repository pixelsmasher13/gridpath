// A server component: the spreadsheet is a client component inside it.
import { Sheet } from "./Sheet";

export default function Page() {
  return (
    <main style={{ display: "flex", flexDirection: "column", height: "100vh" }}>
      <header style={{ padding: "8px 12px", borderBottom: "1px solid #ddd" }}>
        <strong>GridPath in Next.js</strong>
      </header>
      <div style={{ flex: 1, minHeight: 0 }}>
        <Sheet />
      </div>
    </main>
  );
}
