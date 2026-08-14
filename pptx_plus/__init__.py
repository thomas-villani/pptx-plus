"""pptx_plus — OPC-level extensions for python-pptx.

See ``SPEC.md`` at the project root for the public API contract.
"""

from pptx_plus.core import PptxPlusError
from pptx_plus.core._compat import check_upstream_surface

# Fail at import with a message naming the missing attribute, rather than deep
# inside a clone with an AttributeError nobody can act on. SPEC §14.2.
check_upstream_surface()

__all__ = ["PptxPlusError"]
__version__ = "0.1.0"
