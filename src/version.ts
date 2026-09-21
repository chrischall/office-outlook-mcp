/**
 * The single source of the package version.
 *
 * Everything that needs a version imports from here, so release-please has one
 * file to rewrite and there is no second marker to forget. Registered in
 * `release-please-config.json` under `extra-files`.
 *
 * Note for anyone editing the comments in this file: `versionSyncTest` flags
 * every line containing the marker string, so it must appear on the export
 * line below and nowhere else — not even in prose.
 */
export const PACKAGE_NAME = 'office-outlook-mcp';

export const VERSION = '0.1.2'; // x-release-please-version
