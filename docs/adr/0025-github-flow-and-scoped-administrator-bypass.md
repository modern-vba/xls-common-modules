# ADR 0025: Use GitHub Flow with Scoped Administrator Bypass

- Status: Accepted
- Date: 2026-08-30

## Context

The repository is maintained primarily by one maintainer but also accepts pull requests from other contributors. Requiring one approval provides a useful review gate for ordinary contributions, while requiring the maintainer to approve their own pull request would be both impossible in GitHub and procedurally meaningless. Some explicitly requested maintenance and consecutive issue work also needs a direct, auditable path to `main` without turning broad GitHub write access into implicit merge authority.

## Decision

Use GitHub Flow with `main` as the only permanent integration branch and one short-lived branch per issue or maintainer-requested work item.

Ordinary pull requests require one approval, use squash merge only, and are never self-approved. Non-maintainer pull requests always follow this route. A specifically identified maintainer-authored pull request may use administrator bypass only after the maintainer explicitly instructs that pull request to be merged. Repository Administrator bypass is configured as pull-request-only.

Pull-request-less direct integration requires separate, explicit authorization scoped to an identified work item or finite issue set. Only the configured Organization Administrator bypass supports this route. Direct integration preserves linear history, contains one logical work-item commit, and never force-pushes `main`.

Commit, push, origin synchronization, issue and Project operations, closing instructions, and general GitHub write access do not imply integration authorization. Administrator role possession alone does not authorize bypass.

Rules preventing deletion and force-pushes on `main` and requiring linear history have no bypass. GitHub may automatically delete a pull request head branch after a successful merge. Every integration route retains required checks, acceptance-criteria verification on the current `origin/main`, and an issue verification note that records the route, main commit SHA, results, and any bypass authorization.

## Consequences

- External contributions retain a meaningful one-approval review gate.
- Maintainer-authored work does not require mechanical self-approval, but bypass remains an explicit, work-specific decision with an audit trail.
- GitHub write permission and integration permission remain separate concepts.
- Direct integration can support explicitly authorized consecutive work without weakening the normal pull request policy.
- `main` remains linear and protected against deletion and force-pushes regardless of administrator status.
- Issues reach `Done` only after the integrated state, rather than an unmerged branch or pull request, has been verified.
