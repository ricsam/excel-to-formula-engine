# Changesets

This folder holds pending changesets. Add one with `bunx changeset`, describing
the change and whether it is a patch, minor or major.

To release: run `bunx changeset version` to fold pending changesets into
`package.json` and `CHANGELOG.md`, commit the result, and push to `main`. The
publish workflow builds, tests and publishes to npm via trusted publishing.
