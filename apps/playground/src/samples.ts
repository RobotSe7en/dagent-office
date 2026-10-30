import { parseContract } from "@dagent-office/core";
import type { DocumentPackage, EditProposal } from "@dagent-office/core";
async function sample(path: string): Promise<unknown> {
  const response = await fetch(import.meta.env.BASE_URL + path);
  if (!response.ok) throw new Error("协作样例读取失败。");
  return response.json();
}
export async function reportFiles(): Promise<DocumentPackage> {
  const entries = await sample("report-demo.json") as [string, string][];
  const encoder = new TextEncoder();
  return new Map(entries.map(([path, text]) => [path, encoder.encode(text)]));
}
export async function sampleProposal(): Promise<EditProposal> {
  return parseContract("edit-proposal", await sample("report.proposal.json"));
}
