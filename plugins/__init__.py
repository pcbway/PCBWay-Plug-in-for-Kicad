try:
    from .plugin import PCBWayPlugin
    plugin = PCBWayPlugin()
    plugin.register()
except Exception as e:
    print(e)
    import logging
    root = logging.getLogger()
    root.debug(repr(e))
