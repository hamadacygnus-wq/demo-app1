/** Excel 保存先の指定方法（いずれか1つ） */
export function hasExcelTargetConfig(): boolean {
  return Boolean(
    (process.env.GRAPH_DRIVE_ID && process.env.GRAPH_ITEM_ID) ||
      process.env.GRAPH_SHARE_URL?.trim() ||
      process.env.GRAPH_SITE_URL?.trim() ||
      (process.env.GRAPH_USER_UPN?.trim() && process.env.GRAPH_FILE_PATH?.trim()) ||
      (process.env.GRAPH_SITE_ID?.trim() && process.env.GRAPH_FILE_PATH?.trim()) ||
      (process.env.GRAPH_SITE_HOST?.trim() &&
        process.env.GRAPH_SITE_PATH?.trim() &&
        process.env.GRAPH_FILE_PATH?.trim()),
  );
}

export function isGraphConfigured(): boolean {
  return Boolean(
    process.env.AZURE_TENANT_ID &&
      process.env.AZURE_CLIENT_ID &&
      process.env.AZURE_CLIENT_SECRET &&
      hasExcelTargetConfig(),
  );
}
