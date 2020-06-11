class PlatformException(Exception):
    """Custom Exception"""
    def __init__(self, original_exception, custom_message):
        super(PlatformException, self).__init__(f"{custom_message}:\n{original_exception}")
        self.original_exception = original_exception
        self.custom_message = custom_message
