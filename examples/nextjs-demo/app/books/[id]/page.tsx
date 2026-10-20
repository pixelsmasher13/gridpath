// One workbook, by its id: /books/<id>.
import { Workbooks } from "../../Workbooks";

export default async function Page({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  return <Workbooks id={id} />;
}
