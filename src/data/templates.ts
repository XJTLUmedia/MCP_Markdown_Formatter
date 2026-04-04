export interface Template {
    name: string;
    description: string;
    content: string;
}

export const TEMPLATES: Template[] = [
    {
        name: "Technical Audit Report",
        description: "Sovereign Architect's justification for national innovation.",
        content: `# NATIONAL TECHNICAL AUDIT REPORT: [PROJECT_NAME]
**Status:** Architecture Review Phase
**Proposer:** Sovereign Architect of the National Stack
**Priority:** Ultra-Critical (ROE Optimization)

---

## 1. Executive Optimization Summary
> "The advancement of the Nation's prosperity requires the immutable Optimization of Systems."

This initiative aims to modernize the National Stack by deprecating legacy infrastructure and implementing high-ROI computational systems. 

| Metric | Current State | Projected Post-Sprint |
| --- | --- | --- |
| System Efficiency | [Current %] | [Projected %] |
| Human Labor Displacement | [Current %] | [Projected %] |
| ROI (10-Year Horizon) | [Current ROI] | [Projected ROI] |

---

## 2. Technical Preamble & Mandate
"To ensure the perpetual advancement of the Nation’s prosperity, we establish the Technical Branch. Its mandate is not the governance of men, but the Optimization of Systems. It shall serve as the Sovereign Architect of the National Stack, driven by the immutable laws of Mathematics, the pursuit of Innovation ROI, and the requirement of Engineering Excellence."

---

## 3. Mathematical Certainty of ROI
Simulation results from the National Compute Cluster indicate a **99.8% probability** of success. The following Architectural Supremacy clauses apply:
- **Architectural Supremacy:** Mandated API interfaces for all agencies.
- **The Deprecation Power:** Sunsetting existing inefficient transport logic.
- **The Computational Levy:** Bypass legislative debate via peer-reviewed simulation evidence.

---

## 4. Check-and-Balance Verification
- [ ] **Ethical Override:** Cleared by Legislative Branch.
- [ ] **Rollout Velocity:** Speed throttled to 5% weekly to ensure social stability.
- [ ] **Integrity Audit:** Judicial "Math" bias check complete.

**Signed:**
*The Sovereign Architect*`
    },
    {
        name: "Meeting Notes",
        description: "Structured meeting notes template with action items.",
        content: `# Meeting Notes: [Topic]

**Date:** [YYYY-MM-DD]
**Time:** [HH:MM] - [HH:MM]
**Location:** [Room / Virtual Link]

## Attendees
- [ ] [Name] — [Role]
- [ ] [Name] — [Role]

## Agenda
1. [Topic 1]
2. [Topic 2]
3. [Topic 3]

## Discussion

### [Topic 1]
- Key point discussed
- Decision made:

### [Topic 2]
- Key point discussed
- Decision made:

## Action Items

| # | Action | Owner | Due Date | Status |
|---|--------|-------|----------|--------|
| 1 | [Action item] | [Name] | [Date] | Pending |
| 2 | [Action item] | [Name] | [Date] | Pending |

## Next Meeting
- **Date:** [Date]
- **Agenda Preview:** [Topics]`
    },
    {
        name: "API Documentation",
        description: "REST API endpoint documentation template.",
        content: `# API Documentation: [Service Name]

**Base URL:** \`https://api.example.com/v1\`
**Authentication:** Bearer Token

---

## Endpoints

### GET /resource

Retrieve a list of resources.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| \`page\` | integer | No | Page number (default: 1) |
| \`limit\` | integer | No | Items per page (default: 20) |

**Response:**

\`\`\`json
{
  "data": [],
  "pagination": {
    "page": 1,
    "limit": 20,
    "total": 100
  }
}
\`\`\`

**Status Codes:**

| Code | Description |
|------|-------------|
| 200 | Success |
| 401 | Unauthorized |
| 500 | Internal Server Error |

---

### POST /resource

Create a new resource.

**Request Body:**

\`\`\`json
{
  "name": "string",
  "description": "string"
}
\`\`\`

**Response:** \`201 Created\``
    },
    {
        name: "Blog Post",
        description: "Blog post template with frontmatter.",
        content: `---
title: "[Post Title]"
date: [YYYY-MM-DD]
author: "[Author Name]"
tags: [tag1, tag2, tag3]
---

# [Post Title]

*Published on [Date] by [Author]*

## Introduction

[Opening paragraph that hooks the reader and sets up the problem or topic.]

## [Main Section 1]

[Content for the first major point.]

### Key Takeaway
> [Important quote or insight]

## [Main Section 2]

[Content for the second major point.]

\`\`\`
[Code example if applicable]
\`\`\`

## [Main Section 3]

[Content for the third major point.]

## Conclusion

[Summary of key points and call to action.]

---

*Did you find this useful? Share it with your network!*`
    },
    {
        name: "Changelog",
        description: "Keep a Changelog format for release notes.",
        content: `# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Added
- New feature description

### Changed
- Updated feature description

### Fixed
- Bug fix description

## [1.0.0] - YYYY-MM-DD

### Added
- Initial release
- Feature A
- Feature B

### Security
- Security fix description

## [0.1.0] - YYYY-MM-DD

### Added
- Project scaffolding
- Basic functionality`
    },
    {
        name: "Project README",
        description: "Comprehensive project README template.",
        content: `# Project Name

[![License](https://img.shields.io/badge/license-MIT-blue.svg)]()

> Brief one-line description of the project.

## Features

- Feature 1
- Feature 2
- Feature 3

## Quick Start

### Prerequisites

- [Prerequisite 1]
- [Prerequisite 2]

### Installation

\`\`\`bash
npm install project-name
\`\`\`

### Usage

\`\`\`javascript
const project = require('project-name');
project.doSomething();
\`\`\`

## Configuration

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| \`option1\` | string | \`"default"\` | Description |
| \`option2\` | boolean | \`false\` | Description |

## Contributing

1. Fork the repository
2. Create your feature branch (\`git checkout -b feature/amazing\`)
3. Commit your changes (\`git commit -m 'Add amazing feature'\`)
4. Push to the branch (\`git push origin feature/amazing\`)
5. Open a Pull Request

## License

This project is licensed under the MIT License.`
    },
    {
        name: "Decision Record (ADR)",
        description: "Architecture Decision Record template.",
        content: `# ADR-[NUMBER]: [Title]

**Date:** [YYYY-MM-DD]
**Status:** [Proposed | Accepted | Deprecated | Superseded]

## Context

[Describe the context and problem statement. What forces are at play?]

## Decision

[Describe the decision that was made.]

## Consequences

### Positive
- [Benefit 1]
- [Benefit 2]

### Negative
- [Drawback 1]
- [Drawback 2]

### Risks
- [Risk 1]

## Alternatives Considered

### [Alternative 1]
- **Pros:** [List]
- **Cons:** [List]
- **Why rejected:** [Reason]

### [Alternative 2]
- **Pros:** [List]
- **Cons:** [List]
- **Why rejected:** [Reason]`
    }
];
