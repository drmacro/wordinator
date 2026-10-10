# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

The Wordinator generates DOCX files from "Simple Word Processing XML" (SWPX, a.k.a. simplewpml). It is a two-stage pipeline:

1. **XML → SWPX**: optional Saxon XSLT (user-supplied, or the bundled `html2docx` / `ditahtml2docx` transforms in `src/main/xsl/`) turns arbitrary input into one or more SWPX documents.
2. **SWPX → DOCX**: Java code using Apache POI (XWPF) writes DOCX files, with styles taken from a Word template (DOTX/DOCX).

## Commands

Maven project (Java 8 source/target; README says Java 9+ needed at runtime because of POI 5).

- Build + tests: `mvn install`
- Tests only: `mvn test`
- Single test class: `mvn test -Dtest=TestDocxGenerator`
- Single test method: `mvn test -Dtest=TestDocxGenerator#methodName`
- Release package (after `mvn install` has produced the `*-jar-with-dependencies.jar`): `ant` in repo root (`build.xml`, default target `package-release`; version from `version.properties`).
- Run CLI: `java -jar target/wordinator-<ver>-jar-with-dependencies.jar -i <input.xml|dir|.swpx> -o <outdir|file.docx> -t <template.dotx> [-x <transform.xsl>] [-k <catalogs>] [-c <chunklevel>] [-d <dpi>]`. A directory input is treated as a set of `*.swpx` files. `src/main/ant/build.xml` (shipped in the distribution) has `html2docx` and `ditahtml2docx` targets.

Test fixtures live in `src/test/resources/` (`simplewp/`, `html/`, `dita/`, `docx/` templates, `xsl/`, `catalog/`). `out/` and `target/` are build/test output.

## Architecture

- `org.wordinator.xml2docx.MakeDocx`: CLI entry (commons-cli). Decides the path: `.swpx` file → `handleSingleSwpxDoc`; directory → `handleDirectory`; any other XML → `transformXml`, which runs the XSLT via Saxon with a `DocxGeneratingOutputUriResolver` so that every `xsl:result-document` emitting SWPX is converted straight to a DOCX (no intermediate SWPX files needed). Only specific CLI options are passed on as XSLT parameters, not arbitrary extra args: `chunklevel` is always passed (default `root`), and `--fit-images` passes `fit-images=true`. To pass a new parameter, add an option in `MakeDocx.buildOptions()` and put it in `xsltParameters` in `handleCommandLine()`. Values are passed as strings, so declare XSLT params as `xs:string`.
- `generator/DocxGenerator.java` (~3600 lines): the core. Walks SWPX with XMLBeans and builds the POI `XWPFDocument` (paragraphs, runs, tables with spans, footnotes, sections/headers/footers, images, bookmarks, hyperlinks, TOC fields, doc properties, MathML). Table helpers: `RowSpanManager`, `TableColumnDefinition(s)`, `Measurement` (unit parsing). `MathMLConverter` uses Microsoft's `MATHML2OOML.XSL` (not bundled; must be supplied by the user).
- `xwpf/model/`: local overrides of POI classes (e.g. `XWPFHeaderFooterPolicy`) for behavior POI doesn't provide.
- `saxon/`: adapters routing Saxon logging/messages to Log4j.
- `src/main/doctypes/simplewpml/`: RELAX NG grammar for SWPX (`simplewpml.rng`, plus `-mathml2` / `-mathml3` variants as top-level grammars when validating equations). Update grammar when adding SWPX elements/attributes.
- `src/main/xsl/html2docx/`: generic HTML5→SWPX. HTML `@class` → Word style mapping is isolated in a separate mode (`get-style-name.xsl`, XSLT 3 map) so users customize mapping without touching structure logic. `ditahtml2docx/` layers DITA-HTML specifics on top.

New SWPX features usually touch: the RNG grammar, `DocxGenerator`, test SWPX in `src/test/resources/simplewp/`, a test in `TestDocxGenerator`, and a note in the README release notes.

## Conventions

- Non-ServiceNow open-source project: use normal git conventions for branch names (`feature/...`, `release/...`, etc.), not the `scratch/{mergeTarget}/...` scheme from the user's global rules.
- Commit messages reference GitHub issues as `#NNN` (e.g. `Fixes #109:`), not story-number prefixes.
- Default branch is `develop`.
