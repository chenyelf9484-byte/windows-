from docreplace.gui import main
import sys

if __name__ == '__main__':
    if len(sys.argv) >= 3 and sys.argv[1] == '--self-test':
        from docreplace.selftest import run
        sys.exit(run(sys.argv[2], '--with-ocr' in sys.argv))
    else:
        main()
