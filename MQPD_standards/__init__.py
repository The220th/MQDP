from .global_vars import DEBUG_ON
from .standardk import standardk_run

'''
all run must return tuple:
0: str = retrun text
1: int = return type (code)
2: what you want 
'''

__all__ = [
    'standardk_run',
    'DEBUG_ON'
]
