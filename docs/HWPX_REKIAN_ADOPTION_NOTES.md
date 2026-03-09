# HWPX-Rekian Adoption Notes

Status: Proposed  
Date: 2026-03-09  
Owner: `hwp-converter-extension`

## 1. Purpose

This note evaluates what we should and should not adopt from:

- `ai-public-peasant/hwpx-rekian`

The goal is not to copy that repository directly.
The goal is to extract useful production ideas for our HWP/HWPX extension architecture.

Reference repository:

- <https://github.com/ai-public-peasant/hwpx-rekian>

## 2. What That Repository Actually Is

`hwpx-rekian` is primarily:

- an AI skill/toolkit for generating and editing HWPX documents
- centered on XML-first manipulation of HWPX packages
- optimized for document automation workflows, especially public-sector style forms
- partially split into:
  - a pure HWPX XML workflow
  - a Windows COM-based Hancom automation workflow

It is **not**:

- a standalone localhost service for app integration
- a general-purpose production document editor shell
- a macOS-ready packaged extension
- a complete session/checkpoint/recovery system

## 3. High-Value Ideas We Should Adopt

## 3.1 `unpack -> edit -> pack -> validate` as a First-Class Workflow

This is one of the strongest ideas in the repo.

The HWPX skill explicitly treats document work as:

- unpack HWPX
- edit XML parts
- re-pack HWPX
- validate ZIP/XML/package invariants

Why we should adopt it:

- it matches HWPX's actual package structure
- it creates a clean foundation for diffing and testing
- it is much safer than blind text rewriting
- it fits our current structure-preserving direction

How we should use it:

- make this an internal worker workflow, not a user-facing primitive
- preserve original package parts unless intentionally changed
- run validation automatically on every write candidate

## 3.2 HWPX Template Skeletons

The skill uses:

- base template skeletons
- overlay templates for document families such as report/minutes/proposal

Why this matters:

- blank document generation should not start from ad hoc XML emission
- a tested skeleton gives us stable package completeness
- overlays are useful for:
  - document presets
  - enterprise templates
  - automated generation

How we should use it:

- maintain a small curated set of canonical HWPX skeletons
- version them
- validate them in CI
- treat them as controlled assets, not dynamic user-authored core logic

## 3.3 Reference-Based Layout Analysis

The repository explicitly supports generating a new document from the layout of a provided reference HWPX.

Why this is useful:

- it aligns with real enterprise usage
- many users do not want "generic templates"; they want "make another one like this file"
- this pattern is highly valuable for AI-assisted document generation

How we should use it:

- not as arbitrary XML cloning
- but as a structured "template analysis" pipeline:
  - inspect style definitions
  - inspect page geometry
  - inspect table patterns
  - inspect reusable header/footer structures
- emit a normalized internal template model

## 3.4 Markdown Extraction Modes

The skill supports text extraction with options like:

- text only
- include tables
- Markdown output

Why this matters:

- it reinforces the idea that Markdown export should be configurable
- users need different outputs for:
  - AI ingestion
  - human-readable publishing
  - maximum fidelity knowledge capture

How we should use it:

- keep at least:
  - `clean`
  - `fidelity`
- allow table and object handling modes
- always surface diagnostics on what was simplified

## 3.5 Validation as a Required Save Step

The skill validates:

- ZIP validity
- required file presence
- `mimetype` correctness
- XML well-formedness

This is directly compatible with our production goals.

How we should use it:

- validation must run on every save candidate
- failed validation must reject the save
- the previous known-good checkpoint must remain intact

## 3.6 Explicit Style and ID Discipline

The repository documents:

- style IDs
- paragraph/character style references
- ID sequencing discipline

Why this matters:

- HWPX corruption often comes from sloppy ID or style handling
- explicit style accounting helps avoid malformed packages

How we should use it:

- build internal ID allocators
- build style registries
- avoid ad hoc numeric ID creation across the codebase

## 4. Ideas We Should Adopt Only in Modified Form

## 4.1 XML-First Editing

This is powerful, but we should not take it literally as the top-level product model.

What to keep:

- XML part awareness
- package-level operations
- deterministic serialization

What to change:

- our product should edit through a structured internal AST
- direct XML editing should be an implementation detail
- unknown nodes should be preserved pass-through

Decision:

- adopt XML-first internally at the package layer
- do not expose "edit raw section0.xml" as the primary app workflow

## 4.2 Template-Specific Style Maps

The repository includes many hard-coded style IDs and layout conventions for specific document types.

What to keep:

- the idea of explicit style maps
- the usefulness of domain templates

What to change:

- our core should not depend on a fixed set of report/proposal/public-office presets
- those should live as optional template packs on top of the engine

Decision:

- use a template/plugin layer, not a hard-coded product core

## 4.3 Markdown Export Through XML/Text Extraction

Useful as a concept, but not sufficient as a production-quality final design.

What to keep:

- table-aware extraction
- structured text extraction

What to change:

- our production Markdown export should go through:
  - authoritative extraction
  - internal structured model
  - HTML/assets/metadata intermediary
  - Markdown emission

Decision:

- treat `hwpx-rekian` style extraction as inspiration, not final architecture

## 5. Things We Should Not Adopt

## 5.1 Windows COM as a Primary Product Dependency

The repo's `hwp-com-writer` flow relies on:

- Windows
- installed Hancom desktop software
- COM automation

This is not suitable as our primary product runtime because:

- we need macOS support
- we need deterministic packaging
- we need app-bundled or app-managed runtimes
- desktop automation is fragile under scale and sandbox constraints

Decision:

- do not adopt COM as the core architecture
- at most, use it as an internal tooling fallback for Windows-only operations

## 5.2 AI-Written Raw XML as an Authoritative Save Path

The skill workflow is appropriate for agent-assisted generation, but not for direct production saves by end users.

Reasons:

- it is too easy to generate malformed XML
- style and relationship drift become hard to reason about
- package invariants are easy to break
- recovery semantics are weaker

Decision:

- never let freeform AI-generated XML become the primary save mechanism
- AI may propose content or structure, but the engine must normalize and validate it before write

## 5.3 Skill-Style Shell Scripts as Runtime Integration

The repository is organized around skills, scripts, and examples.

That is useful for experimentation, not for production app embedding.

Decision:

- do not adopt its shell-script driven integration model
- our integration boundary remains:
  - versioned service API
  - managed runtime
  - host-controlled sessions

## 5.4 Product Logic Coupled to Specific Document Families

The repository leans heavily into public-sector document families.

That is a feature pack, not a universal engine.

Decision:

- do not let our HWP engine architecture depend on one domain's layout assumptions
- domain templates should be optional packs layered above the engine

## 6. Concrete Things We Should Add to Our Roadmap

Based on this review, we should explicitly add the following work items.

## 6.1 HWPX Package Toolkit

Add an internal package toolkit with:

- unpack
- repack
- pretty-print for debug output
- package validation
- part manifest inspection

## 6.2 Canonical Skeleton Assets

Add versioned canonical skeletons for:

- blank document
- table-heavy document
- header/footer document

Later:

- enterprise template packs

## 6.3 Reference-Template Analyzer

Add a tool that converts a sample HWPX into:

- style inventory
- page geometry summary
- table layout summary
- reusable template metadata

## 6.4 Export Diagnostics

Add diagnostics for Markdown export:

- unsupported object count
- table downgrade count
- font substitution notes
- preserved assets count
- metadata sidecar path

## 6.5 Validation Gate

Make save impossible unless:

- package validation passes
- reopen verification passes
- checkpoint creation succeeds

## 7. How This Fits Our Existing Architecture

This review reinforces our current direction:

- `HWPX` remains the canonical editable format
- `.hwp` remains import-first
- production save stays structure-preserving
- the extension remains a host-integrated service, not a loose script pack

What changes:

- we should strengthen the internal HWPX tooling layer
- we should add explicit skeleton/template and reference-analysis capabilities
- we should improve Markdown export modes and diagnostics

## 8. Final Recommendation

From `hwpx-rekian`, we should adopt:

- package-oriented HWPX thinking
- unpack/edit/pack/validate discipline
- skeleton/template assets
- reference-based generation ideas
- configurable Markdown extraction modes

We should not adopt:

- Windows COM as the primary runtime
- freeform XML authoring as the main save model
- skill-shell based integration as product architecture
- domain-specific public-office template assumptions as engine core

In short:

- **adopt the package discipline**
- **do not adopt the runtime model**

## 9. Source Notes

Primary source used for this review:

- `hwpx` skill description and workflow:
  - <https://raw.githubusercontent.com/ai-public-peasant/hwpx-rekian/master/hwpx/SKILL.md>
- `hwp-com-writer` skill description:
  - <https://raw.githubusercontent.com/ai-public-peasant/hwpx-rekian/master/hwp-com-writer/SKILL.md>

