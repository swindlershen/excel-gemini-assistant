#!/usr/bin/env bash
# Builds and publishes dist/ to the gh-pages branch as a normal (non-force) commit,
# so GitHub Pages stays enabled across deployments.
set -euo pipefail
cd "$(dirname "$0")/.."
REPO_URL=$(git remote get-url origin)

npm run build
touch dist/.nojekyll

WORK=$(mktemp -d)
git clone -q --branch gh-pages --single-branch "$REPO_URL" "$WORK" 2>/dev/null || {
  git init -q -b gh-pages "$WORK"
  git -C "$WORK" remote add origin "$REPO_URL"
}
find "$WORK" -mindepth 1 -maxdepth 1 ! -name .git -exec rm -rf {} +
cp -r dist/. "$WORK"/
git -C "$WORK" add -A
if git -C "$WORK" diff --cached --quiet; then
  echo "No changes to deploy."
else
  git -C "$WORK" commit -qm "Deploy $(git rev-parse --short HEAD)"
  git -C "$WORK" push -q origin gh-pages
  echo "Deployed."
fi
rm -rf "$WORK"
