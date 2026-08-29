# Cloudflare deployment

This version is adapted from the original Express/local-file app for Cloudflare Workers.

- Static pages are served from `public/` through the Workers Assets binding.
- API routes run in `src/index.js`.
- Persistent state is stored in the D1 database `mvu-report-db`.
- The D1 binding is named `DB` in `wrangler.toml`.
- The database table `app_state` is created automatically on first API request.

## Deploy from the existing GitHub-connected Worker

1. Commit and push the new `src/index.js` and `wrangler.toml` to the production `main` branch.
2. Cloudflare should start a new build automatically.
3. The deploy command should be `npx wrangler deploy`.
4. Open the Worker URL and test `/api/status` and the upload page.

The old `server.js` remains in the repository for local/legacy use, but Cloudflare deploys `src/index.js` because that is the `main` entry in `wrangler.toml`.


## D1 large-file fix
Village Mapping and Week Off data are stored in normalized D1 tables in batches instead of one large JSON row. This prevents SQLITE_TOOBIG for large Excel masters. Hard Reset clears the normalized tables and the legacy app_state table.
