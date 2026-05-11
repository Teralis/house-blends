# Consulting Automations (Public)
This repo holds client automations that are safe to be public — no credentials, no sensitive business logic. Each top-level folder is one client-automation.

## Working style
- I'm not a developer. Prefer plain scripts over frameworks.
- Vibe-coded, not engineered. Skip abstractions, base classes, type hierarchies.
- Keep dependencies minimal.

## Repo conventions
- Each project folder is self-contained: own requirements.txt, own README,
  own CLAUDE.md if needed.
- Workflows live in /.github/workflows/ at the root, named <client>-<job>.yml.
- Project-specific context, quirks, and "don'ts" go in that project's CLAUDE.md.

## What's where
- Active: smc-submittal-report
- Paused: none
- Dead: none
