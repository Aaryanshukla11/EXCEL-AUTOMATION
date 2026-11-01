# Autonomous Excel Intelligence Platform

An enterprise-grade, generic AI automation platform that understands, plans, consolidates, executes, and validates complex Excel workflows without hardcoded templates or brittle rules.

## Core Philosophy
1. **AI Thinks and Plans**: Uses LLMs and reasoning models for structural discovery, semantic mapping, and execution plan formulation.
2. **Python Deterministically Executes**: Pure deterministic Python scripts using OpenPyXL and XML trees handle all cell modifications, format preservation, and formula protection.
3. **Multi-Signal Entity Matching**: Combines strong identifiers (e.g. Employee ID, Account ID), fuzzy name similarity, and contextual organizational signals with confidence scoring.
4. **Validation Verifies**: Independent pre- and post-write verification checks formula integrity, row/column boundaries, and checksum consistency.
5. **Memory Learns**: Experience-based memory records user-approved decisions and rules without fine-tuning or hardcoding.
6. **Audit & Rollback**: Complete step-by-step audit logging and non-destructive backups for instant, single-click rollback.

## Tech Stack
- **Frontend**: React, TypeScript, Tailwind CSS, Vite.
- **Backend**: Python 3.12/3.13, FastAPI, Pydantic v2, SQLAlchemy.
- **Excel Processing**: OpenPyXL, Pandas, lxml.
- **Matching & Similarity**: RapidFuzz, Levenshtein, custom multi-signal scoring.
- **Database**: SQLite (dev) / PostgreSQL (production).

## Project Structure
Refer to [ARCHITECTURE.md](file:///C:/Users/Aaryan%20shukla/OneDrive/Desktop/New%20folder/autonomous-excel-intelligence/ARCHITECTURE.md) for full architectural documentation.

