/** Browser and worker children inherit runtime/font configuration, not host credentials. */
export function conversionEnvironment(): Record<string, string> {
  const result: Record<string,string> = {};
  for (const key of ["PATH", "HOME", "USER", "LOGNAME", "LANG", "LC_ALL", "FONTCONFIG_FILE", "FONTCONFIG_PATH", "XDG_CACHE_HOME", "XDG_CONFIG_HOME", "XDG_RUNTIME_DIR", "TMPDIR", "DISPLAY", "WAYLAND_DISPLAY", "LD_LIBRARY_PATH", "PLAYWRIGHT_BROWSERS_PATH"])
    if (process.env[key] !== undefined) result[key] = process.env[key];
  return result;
}
