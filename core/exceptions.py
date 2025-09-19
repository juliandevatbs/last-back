# Custom exceptions for project

class NoDataError(Exception):

    # No data received to write
    pass


class KeyNotFound(Exception):

    def __init__(self, key: str):

        # No key in the dict
        message = f"Error key {key} not found"
        super().__init__(message)
        self.key = key



    pass


class ServerClientException(Exception):

    pass
