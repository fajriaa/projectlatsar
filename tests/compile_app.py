"""
Test module to verify that the main application file compiles without errors.
"""

import compileall
import traceback
import sys


def test_app_compilation():
    """Compile app.py and report compilation status."""
    try:
        compileall.compile_file('app.py', doraise=True)
        print('✓ COMPILE_OK: app.py compiled successfully')
        return True
    except SyntaxError as e:
        print(f'✗ COMPILE_ERROR: Syntax error in app.py')
        traceback.print_exc()
        return False
    except Exception as e:
        print(f'✗ COMPILE_ERROR: {type(e).__name__}: {e}')
        traceback.print_exc()
        return False


if __name__ == '__main__':
    success = test_app_compilation()
    sys.exit(0 if success else 1)
