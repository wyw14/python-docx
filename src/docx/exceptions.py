"""Exceptions used with python-docx.

The base exception class is PythonDocxError.
"""


class PythonDocxError(Exception):
    """Generic error class."""


class ImageReferenceError(PythonDocxError):
    """Raised when an embedded-image reference in document XML cannot be resolved to a
    valid image part.

    This occurs for example when an ``a:blip`` element's ``r:embed`` attribute names
    a relationship that is missing from the story part's relationship collection, or
    when that relationship does not target an image part.
    """


class InvalidSpanError(PythonDocxError):
    """Raised when an invalid merge region is specified in a request to merge table
    cells."""


class InvalidXmlError(PythonDocxError):
    """Raised when invalid XML is encountered, such as on attempt to access a missing
    required child element."""
