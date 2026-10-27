#!/usr/bin/env python3
"""Editorial preflight for structured scientific manuscripts.

Detects obvious draft traces, unsupported promotional wording and example data.
It does not fact-check numerical results, literature or scientific derivations.
"""
from __future__ import annotations
import argparse
import json
import re
import sys
from pathlib import Path

from crossrefs import make_anchors, validate_refs

CLAIM = re.compile(r'\b(significantly|dramatically|exceptional(?:ly)?|perfect(?:ly)?|unprecedented)\b', re.I)
PLACEHOLDER = re.compile(r'\b(?:CITATION NEEDED|REFERENCE NEEDED|INSERT FIGURE|TBD|TODO)\b|\[\?\]', re.I)
DRAFT_TRACE = re.compile(r'\b(?:MERGEFORMAT|this section is assumed to be replaced|this section is assumed to be rewritten|Equation Chapter)\b', re.I)
SAMPLE_TRACE = re.compile(
    r'(?:formatting demonstration only|formatting examples? only|not research data|'
    r'synthetic data|all values are synthetic|replace with a verified bibliographic record|'
    r'examples?\.org|Example University|A\. Researcher|B\. Researcher)', re.I)


def lint(spec, submission=False):
    """Return issues; submission=True upgrades incomplete/synthetic drafts to errors."""
    issues = []
    def issue(level, location, message, excerpt=None):
        row = {'level': level, 'location': location, 'message': message}
        if excerpt:
            row['excerpt'] = excerpt[:160]
        issues.append(row)

    try:
        refs, _ = make_anchors(spec)
        validate_refs(spec, refs)
    except (ValueError, TypeError) as exc:
        issue('error', 'structure', str(exc))

    text_sources = [('title', spec.get('title', '')), ('abstract', spec.get('abstract', ''))]
    for i, name in enumerate(spec.get('authors', [])):
        text_sources.append((f'authors[{i}]', name))
    for i, name in enumerate(spec.get('affiliations', [])):
        text_sources.append((f'affiliations[{i}]', name))
    text_sources.append(('corresponding_author', spec.get('corresponding_author', '')))
    for i, section in enumerate(spec.get('sections', [])):
        text_sources.append((f'sections[{i}].heading', section.get('heading', '')))
        for j, block in enumerate(section.get('blocks', [])):
            location = f'sections[{i}].blocks[{j}]'
            for key in ('heading', 'text', 'caption', 'note'):
                if key in block:
                    text_sources.append((f'{location}.{key}', block[key]))
            for k, segment in enumerate(block.get('segments', [])):
                if isinstance(segment, dict) and 'text' in segment:
                    text_sources.append((f'{location}.segments[{k}]', segment['text']))
    for i, ref in enumerate(spec.get('references', [])):
        text_sources.append((f'references[{i}]', ref))

    for location, text in text_sources:
        if not isinstance(text, str):
            issue('error', location, 'Expected a string')
            continue
        if CLAIM.search(text):
            issue('warning', location, 'Strong claim: verify quantitative evidence', text)
        if PLACEHOLDER.search(text):
            issue('error', location, 'Unresolved placeholder', text)
        if DRAFT_TRACE.search(text):
            issue('error', location, 'Legacy draft/Word field artifact', text)
        if SAMPLE_TRACE.search(text):
            issue('error' if submission else 'warning', location,
                  'Synthetic/example content: never submit as scientific evidence', text)

    if spec.get('sample_data') is True:
        issue('error' if submission else 'warning', 'sample_data',
              'Illustrative example flagged as sample_data=true')
    if not spec.get('abstract'):
        issue('error' if submission else 'warning', 'abstract', 'Missing abstract')
    if not spec.get('references'):
        issue('error' if submission else 'warning', 'references', 'Missing bibliography')
    if not spec.get('sections'):
        issue('error', 'sections', 'No sections')
    if submission and not spec.get('authors'):
        issue('error', 'authors', 'Missing authors for submission draft')
    return issues


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('spec')
    ap.add_argument('--submission', action='store_true',
                    help='Block synthetic/unfinished manuscript inputs before submission')
    args = ap.parse_args()
    spec = json.loads(Path(args.spec).read_text(encoding='utf-8'))
    issues = lint(spec, submission=args.submission)
    print(json.dumps({
        'submission_mode': args.submission,
        'errors': sum(x['level'] == 'error' for x in issues),
        'warnings': sum(x['level'] == 'warning' for x in issues),
        'issues': issues,
        'note': 'This tool cannot verify scientific validity or bibliographic accuracy.'
    }, ensure_ascii=False, indent=2))
    return 2 if any(x['level'] == 'error' for x in issues) else 0


if __name__ == '__main__':
    sys.exit(main())
