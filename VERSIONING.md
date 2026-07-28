# Versioning Policy

ProAV Shōko follows **Semantic Versioning (SemVer 2.0.0)**.

```
MAJOR.MINOR.PATCH[-PRERELEASE][+BUILD]
```

## Components

| Part | When to increment | Examples |
|------|------------------|----------|
| **MAJOR** | Breaking changes (removed features, incompatible formats, API changes) | `1.0.0` → `2.0.0` |
| **MINOR** | New features, backward-compatible | `1.0.0` → `1.1.0` |
| **PATCH** | Bug fixes, performance, UI polish | `1.1.0` → `1.1.1` |

## Pre-release Stages

| Stage | Format | Description |
|-------|--------|-------------|
| **Alpha** | `0.1.0-alpha.1` | Incomplete features, bugs expected, internal testing |
| **Beta** | `0.1.0-beta.1` | Feature-complete, bug fixes only, public testers |
| **RC** | `0.1.0-rc.1` | Only critical fixes, candidate for final release |

## Release Workflow

```
Dev → alpha.1 → alpha.n → beta.1 → beta.n → rc.1 → 0.1.0 → 0.1.1 → 0.2.0 → 1.0.0
```

## Build Metadata

Auto-appended by CI: `0.4.0-alpha.3+278` (build number for diagnostics only).

## Executable Naming

```
ProAV Shoko 0.1.0-beta.2.exe
```

Use SemVer — avoid custom formats like `ProAV Shoko.0.0.b0.01.exe`.

## Automatic Versioning

CI auto-increments: pre-release number, build number, filenames, tags.  
Manual only: **MAJOR**, **MINOR**, **PATCH** (intentional milestones).

## Philosophy

- Predictable versions → users understand release maturity at a glance.
- Developers identify builds quickly.
- Automated builds never require manual renaming.
