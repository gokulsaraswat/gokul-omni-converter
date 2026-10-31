"""One module per screen.

Existing screens expose ``build(app)`` and assign their widgets onto the app so
the handler methods in ``app.py`` keep working unchanged. The two screens added
in 2.4.0 (archive, rename_organize) are self-contained page classes that own
both their widgets and their handlers.
"""
