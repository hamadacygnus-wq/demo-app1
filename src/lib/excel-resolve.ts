import { graphFetch } from "./graph-auth";

export type WorkbookTarget = { driveId: string; itemId: string };

/** 共有リンクを Graph の shareId 形式に変換 */
export function encodeShareUrl(url: string): string {
  const base64 = Buffer.from(url, "utf8").toString("base64");
  const unpadded = base64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  return `u!${unpadded}`;
}

function encodeDrivePath(filePath: string): string {
  return filePath
    .replace(/^\//, "")
    .split("/")
    .map((segment) => encodeURIComponent(segment))
    .join("/");
}

async function itemToTarget(item: {
  id: string;
  parentReference?: { driveId?: string };
}): Promise<WorkbookTarget> {
  const driveId = item.parentReference?.driveId;
  if (!driveId) {
    throw new Error("ファイルのドライブIDを自動取得できませんでした");
  }
  return { driveId, itemId: item.id };
}

async function resolveByShareUrl(shareUrl: string): Promise<WorkbookTarget> {
  const shareId = encodeShareUrl(shareUrl.trim());
  const res = await graphFetch(`/shares/${shareId}/driveItem`);
  if (!res.ok) {
    const text = await res.text();
    throw new Error(
      `共有リンクからファイルを取得できませんでした: ${res.status} ${text}\n` +
        "リンクの共有範囲と、アプリの Files.ReadWrite.All 権限を確認してください。",
    );
  }
  return itemToTarget((await res.json()) as { id: string; parentReference?: { driveId?: string } });
}

async function resolveByUserPath(upn: string, filePath: string): Promise<WorkbookTarget> {
  const encoded = encodeDrivePath(filePath);
  const res = await graphFetch(`/users/${encodeURIComponent(upn)}/drive/root:/${encoded}`);
  if (!res.ok) {
    const text = await res.text();
    throw new Error(
      `OneDrive からファイルを取得できませんでした: ${res.status} ${text}\n` +
        "GRAPH_USER_UPN（メールアドレス）と GRAPH_FILE_PATH を確認してください。",
    );
  }
  return itemToTarget((await res.json()) as { id: string; parentReference?: { driveId?: string } });
}

async function resolveSiteId(hostname: string, sitePath: string): Promise<string> {
  const normalizedPath = sitePath.startsWith("/") ? sitePath : `/${sitePath}`;
  const res = await graphFetch(`/sites/${hostname}:${normalizedPath}`);
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`SharePoint サイトの取得に失敗しました: ${res.status} ${text}`);
  }
  const json = (await res.json()) as { id: string };
  return json.id;
}

async function resolveBySiteAndPath(siteId: string, filePath: string): Promise<WorkbookTarget> {
  const encoded = encodeDrivePath(filePath);
  const res = await graphFetch(`/sites/${siteId}/drive/root:/${encoded}`);
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`SharePoint 上のファイル取得に失敗しました: ${res.status} ${text}`);
  }
  return itemToTarget((await res.json()) as { id: string; parentReference?: { driveId?: string } });
}

/**
 * ブラウザのアドレスバーに表示される SharePoint / OneDrive の URL からファイルを特定
 * 例: https://tenant.sharepoint.com/sites/MySite/Shared Documents/joy/食事記録.xlsx
 */
export function parseSharePointFileUrl(pageUrl: string): {
  hostname: string;
  sitePath: string;
  filePath: string;
} | null {
  let url: URL;
  try {
    url = new URL(pageUrl.trim());
  } catch {
    return null;
  }

  const hostname = url.hostname;
  const pathname = decodeURIComponent(url.pathname);

  // 個人用 OneDrive: /personal/user_domain_com/Documents/...
  const personalMatch = pathname.match(/^\/personal\/([^/]+)\/(.+)$/i);
  if (personalMatch && hostname.includes("-my.sharepoint.com")) {
    const rest = personalMatch[2];
    const docIdx = rest.search(/\/Documents\//i);
    if (docIdx >= 0) {
      const afterDocuments = rest.slice(docIdx + "/Documents/".length);
      const filePath = afterDocuments.startsWith("/")
        ? `Documents${afterDocuments}`
        : `Documents/${afterDocuments}`;
      return {
        hostname,
        sitePath: `/personal/${personalMatch[1]}`,
        filePath,
      };
    }
  }

  // チームサイト: /sites/SiteName/Shared Documents/... または /Shared Documents/...
  const sharedIdx = pathname.search(/\/(Shared Documents|共有ドキュメント)\//i);
  if (sharedIdx < 0) return null;

  const filePath = pathname
    .slice(sharedIdx + 1)
    .replace(/^Shared Documents\//i, "Shared Documents/")
    .replace(/^共有ドキュメント\//, "Shared Documents/");

  const beforeShared = pathname.slice(0, sharedIdx);
  const siteMatch = beforeShared.match(/(\/sites\/[^/]+|\/teams\/[^/]+)/i);
  const sitePath = siteMatch ? siteMatch[1] : "";

  if (!sitePath && !hostname.includes("-my.sharepoint.com")) {
    return null;
  }

  return {
    hostname,
    sitePath: sitePath || "/",
    filePath: filePath.replace(/^\//, ""),
  };
}

async function resolveBySiteUrl(pageUrl: string): Promise<WorkbookTarget> {
  const parsed = parseSharePointFileUrl(pageUrl);
  if (!parsed) {
    throw new Error(
      "GRAPH_SITE_URL の形式を解釈できませんでした。SharePoint でファイルを開いたときのアドレスバーの URL をそのまま貼り付けてください。",
    );
  }

  const siteId = await resolveSiteId(parsed.hostname, parsed.sitePath);
  return resolveBySiteAndPath(siteId, parsed.filePath);
}

let targetCache: WorkbookTarget | null = null;

export async function resolveWorkbookTarget(): Promise<WorkbookTarget> {
  if (targetCache) return targetCache;

  // 方式1: ドライブID + ファイルID（上級者向け）
  if (process.env.GRAPH_DRIVE_ID && process.env.GRAPH_ITEM_ID) {
    targetCache = {
      driveId: process.env.GRAPH_DRIVE_ID,
      itemId: process.env.GRAPH_ITEM_ID,
    };
    return targetCache;
  }

  // 方式2: 共有リンク（いちばん簡単）
  if (process.env.GRAPH_SHARE_URL?.trim()) {
    targetCache = await resolveByShareUrl(process.env.GRAPH_SHARE_URL);
    return targetCache;
  }

  // 方式3: ブラウザの URL をそのまま貼る
  if (process.env.GRAPH_SITE_URL?.trim()) {
    targetCache = await resolveBySiteUrl(process.env.GRAPH_SITE_URL);
    return targetCache;
  }

  // 方式4: OneDrive — ユーザーのメール + ファイルパス
  if (process.env.GRAPH_USER_UPN?.trim() && process.env.GRAPH_FILE_PATH?.trim()) {
    targetCache = await resolveByUserPath(
      process.env.GRAPH_USER_UPN.trim(),
      process.env.GRAPH_FILE_PATH.trim(),
    );
    return targetCache;
  }

  const filePath = process.env.GRAPH_FILE_PATH?.trim();
  if (!filePath) {
    throw new Error(
      "Excel の保存先が未設定です。GRAPH_SHARE_URL、GRAPH_SITE_URL、または GRAPH_USER_UPN + GRAPH_FILE_PATH のいずれかを .env.local に設定してください。",
    );
  }

  // 方式5: サイトID + パス
  if (process.env.GRAPH_SITE_ID?.trim()) {
    targetCache = await resolveBySiteAndPath(process.env.GRAPH_SITE_ID.trim(), filePath);
    return targetCache;
  }

  // 方式6: サイトのホスト名 + サイトパス + ファイルパス
  const host = process.env.GRAPH_SITE_HOST?.trim();
  const sitePath = process.env.GRAPH_SITE_PATH?.trim();
  if (host && sitePath) {
    const siteId = await resolveSiteId(host, sitePath);
    targetCache = await resolveBySiteAndPath(siteId, filePath);
    return targetCache;
  }

  throw new Error(
    "Excel の保存先が未設定です。GRAPH_SHARE_URL（共有リンク）の設定を推奨します。",
  );
}
