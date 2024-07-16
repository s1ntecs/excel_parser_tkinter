from functools import wraps
from exceptions import ProcessorNotLoadedError


def requires_processor(func):
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        if not hasattr(self, 'processor'):
            self.handle_error(ProcessorNotLoadedError)
            return
        return func(self, *args, **kwargs)
    return wrapper