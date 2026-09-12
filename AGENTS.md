# Repository instructions

## Education mirrors

When modifying a production Python source file, also create or update its
corresponding mirror under `Education/`. Production files are authoritative;
their mirrors must contain exactly the same executable Python code and behavior,
with only additional educational comments.

Education mirrors are learning artifacts only. Do not import, execute, or use
them as production programs or inputs.

Write cumulative, substantive Python lessons for an experienced JavaScript/
Node.js and VBA developer learning Python. Each lesson block begins once with
the exact marker `# Lesson:Python`. Explain concepts and idioms, including
useful JavaScript comparisons, rather than restating obvious lines. Do not
repeat an introductory lesson already covered by another mirror unless the new
use has a materially different aspect.

Before finishing a change to a mirrored file, verify executable equivalence
with Python's standard-library `ast` module: parse both files and compare
`ast.dump(tree, include_attributes=False)`. This deliberately ignores comments
and formatting while detecting changes to Python syntax or behavior.
