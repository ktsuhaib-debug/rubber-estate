# Rubber Estate

Family app for tracking the two rubber estates: daily tapping, estate work, sheet sales, tapper wages, and PDF statements.

- **Root folder** – the full-screen launcher served by GitHub Pages (icon, manifest, and a page that loads the app).
- **`app/`** – the Apps Script code that runs the app (`Code.gs`, `Index.html`) and `version.json`.

## Updating the app

The app updates itself from the `app/` folder:

1. Push new `app/Code.gs` and `app/Index.html`, set the same version in `APP_VERSION` (Code.gs) and `app/version.json`.
2. In the app, an admin opens **Settings → Account → App updates → Check for updates → Update now**.

The app checks the files, replaces its own code, creates a new version and moves the existing deployment to it, so the link never changes.
