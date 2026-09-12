# Education

Files here are educational mirrors of production source files. Production source
remains authoritative; each mirror has the same executable code with additional
explanatory comments and is never a production input.

Lessons begin with searchable markers such as `Lesson:Python`. JavaScript/Node.js
comparisons are included where helpful. Lessons form a cumulative curriculum, so
they intentionally avoid reteaching concepts already covered elsewhere.

To verify a Python mirror, parse it and its production counterpart with Python's
`ast` module and compare `ast.dump(tree, include_attributes=False)`; comments
are ignored while executable syntax is checked.
