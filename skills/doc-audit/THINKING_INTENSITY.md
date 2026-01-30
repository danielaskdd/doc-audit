# LLM Thinking Intensity Control

This document describes how to control the thinking/reasoning intensity of LLM models when running audits.

## Overview

The audit system now supports controlling how deeply LLMs "think" about each document block:

- **Gemini 3 models**: Use `thinking_level` (minimal, low, medium, high)
- **Gemini 2.5 models**: Use `thinking_budget` (token count, 0 to disable)
- **OpenAI o-series models**: Use `reasoning_effort` (low, medium, high)

## Configuration Methods

### Command Line Arguments (Recommended)

```bash
# For Gemini 3 models (e.g., gemini-3-pro-preview)
python run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --model gemini-3-pro-preview \
  --thinking-level high

# For Gemini 2.5 models (e.g., gemini-2.5-pro)
python run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --model gemini-2.5-pro \
  --thinking-budget 1024

# For OpenAI o-series models (e.g., o1-mini, o3-mini)
python run_audit.py \
  --document blocks.jsonl \
  --rules rules.json \
  --model o3-mini \
  --reasoning-effort medium
```

### Environment Variables

Set these before running the audit:

```bash
# For Gemini 3
export GEMINI_THINKING_LEVEL="high"

# For Gemini 2.5
export GEMINI_THINKING_BUDGET="1024"

# For OpenAI o-series
export OPENAI_REASONING_EFFORT="medium"
```

**Priority**: Command line arguments override environment variables.

## Parameter Details

### Gemini 3: `--thinking-level`

Controls the depth of reasoning for Gemini 3 models.

| Level | Description | Use Case |
|-------|-------------|----------|
| `minimal` | Fastest, basic reasoning | Quick checks, simple rules |
| `low` | Light reasoning | General audits |
| `medium` | Balanced reasoning | Standard audits (recommended) |
| `high` | Deep reasoning | Complex rules, critical documents |

### Gemini 2.5: `--thinking-budget`

Token budget for internal thinking (Gemini 2.5 Pro/Flash).

| Budget | Description | Use Case |
|--------|-------------|----------|
| `0` | Disabled | Fast mode, simple tasks |
| `128` | Minimum (2.5-pro) | Basic reasoning |
| `512` | Moderate | Standard audits |
| `1024` | High | Complex analysis |
| `2048+` | Maximum | Critical documents |

**Note**: Gemini 2.5 Flash thinking is enabled by default; set to 0 to disable for faster response.

### OpenAI: `--reasoning-effort`

Reasoning intensity for o-series models (o1, o3-mini, etc.).

| Level | Description | Use Case |
|-------|-------------|----------|
| `low` | Quick reasoning | Simple rules |
| `medium` | Balanced | Standard audits (recommended) |
| `high` | Deep reasoning | Complex analysis |

## Examples

### Example 1: High-Quality Audit with Gemini 3

```bash
python run_audit.py \
  --document contract_blocks.jsonl \
  --rules legal_rules.json \
  --model gemini-3-pro-preview \
  --thinking-level high \
  --workers 4
```

### Example 2: Fast Audit with Gemini 2.5 Flash

```bash
python run_audit.py \
  --document simple_doc_blocks.jsonl \
  --rules basic_rules.json \
  --model gemini-2.5-flash \
  --thinking-budget 0 \
  --workers 8
```

### Example 3: Balanced OpenAI o-series Audit

```bash
python run_audit.py \
  --document technical_blocks.jsonl \
  --rules tech_rules.json \
  --model o3-mini \
  --reasoning-effort medium \
  --workers 6
```

## Performance Considerations

### Speed vs Quality Trade-offs

| Configuration | Speed | Quality | Cost | Best For |
|---------------|-------|---------|------|----------|
| Minimal/Low/0 | ⚡⚡⚡ | ⭐⭐ | 💰 | Simple rules, quick checks |
| Medium/512 | ⚡⚡ | ⭐⭐⭐ | 💰💰 | Standard audits |
| High/1024+ | ⚡ | ⭐⭐⭐⭐ | 💰💰💰 | Critical documents |

### Recommendations

1. **Start with defaults** (no thinking params) for most audits
2. **Use high thinking** for:
   - Legal/compliance documents
   - Complex technical specifications
   - High-stakes content
3. **Use low/disabled thinking** for:
   - Simple grammar checks
   - Formatting verification
   - Quick sanity checks

## Integration with Workflow

The `workflow.sh` script doesn't directly support thinking parameters yet. To use them:

```bash
# Activate environment
source .claude-work/doc-audit/env.sh

# Run audit with thinking control
python $DOC_AUDIT_SKILL_PATH/scripts/run_audit.py \
  --document document_blocks.jsonl \
  --rules .claude-work/doc-audit/default_rules.json \
  --thinking-level high \
  --output document_manifest.jsonl

# Generate report (same as before)
python $DOC_AUDIT_SKILL_PATH/scripts/generate_report.py \
  document_manifest.jsonl \
  --template .claude-work/doc-audit/report_template.html \
  --output document_audit_report.html
```

## Troubleshooting

### No Effect Observed

- **Check model compatibility**: Only Gemini 2.5+/3+ and OpenAI o-series support these parameters
- **Verify parameter spelling**: `--thinking-level`, `--thinking-budget`, `--reasoning-effort`
- **Look for startup message**: The script prints thinking configuration at startup

### Invalid Parameter Error

```
Error: unrecognized arguments: --thinking-level
```

**Solution**: Make sure you're using the updated `run_audit.py` script.

### Model Not Found

```
Error: Model gemini-3-pro-preview not found
```

**Solution**: Use available models. Check Gemini/OpenAI documentation for current model names.

## See Also

- [Gemini Thinking Documentation](https://ai.google.dev/gemini-api/docs/thinking)
- [OpenAI Reasoning Documentation](https://platform.openai.com/docs/guides/reasoning)
- [run_audit.py Documentation](skills/doc-audit/TOOLS.md)
