"""CustomTkinter presentation layer for GO Omni Pro.

``ui.kit`` holds the widget factories every page builds from; ``ui.pages``
holds one module per screen. Pages never construct raw widgets or name colours
directly -- they ask the kit, which reads the design tokens in ``ui_theme``.
"""
