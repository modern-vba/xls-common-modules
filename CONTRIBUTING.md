# Contributing

This repository uses GitHub Flow. `main` is the only permanent integration branch, and every issue or maintainer-requested work item is developed on one short-lived branch created from the current `origin/main`.

## Pull Requests

- Use a pull request by default and keep it limited to one logical work item.
- Do not base a later issue or work-item branch on an unfinished branch. Start it from the updated `origin/main` only after the earlier work is integrated.
- Reference an issue with `Refs #123`; do not use an automatic closing keyword. The issue is closed only after the integrated change is verified on the current `origin/main`.
- All ordinary pull requests require one approval. Authors must never approve their own pull requests.
- Non-maintainer pull requests always follow this reviewed route.
- Only squash merge is allowed. Use an English Conventional Commit title for the pull request.
- GitHub automatically deletes a successfully merged pull request head branch. Remove local branches only after post-merge verification.

## Maintainer Exceptions

A specifically identified maintainer-authored pull request may use administrator bypass only after the maintainer explicitly instructs that pull request to be merged. The bypass does not waive testing, acceptance criteria, audit notes, squash merge, or post-merge verification.

Pull-request-less direct integration is a separate exception. It requires explicit, scoped authorization for the identified work item or finite issue set. It uses the configured Organization Administrator bypass, preserves linear history, and must update `main` without a force-push. Repository Administrator bypass is limited to pull requests.

Permission to commit, push, synchronize with origin, update GitHub Issues or Projects, close work, or otherwise write to GitHub is not permission to merge or integrate into `main`. Administrator access by itself is not authorization to bypass the normal route.

## Protected Main and Verification

Protection from deletion, force-pushes, and non-linear history is invariant and has no bypass. After either integration route, fetch the current `origin/main` and verify the intended change, required checks, and acceptance criteria there. Before an issue moves to `Done`, its verification note records the integration route, main commit SHA, acceptance result, check results, and any administrator bypass authorization.

The rationale and exact policy are recorded in [ADR 0025](docs/adr/0025-github-flow-and-scoped-administrator-bypass.md).
