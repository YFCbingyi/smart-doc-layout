"""允许通过 python -m word_processor 直接运行。

用法:
    python -m word_processor read <input.docx> [output.json]
    python -m word_processor modify <input.docx> <modify.json> [output.docx]
    python -m word_processor gui
"""
import sys

if len(sys.argv) > 1 and sys.argv[1] == "gui":
    from word_processor.gui import main as gui_main
    gui_main()
else:
    from word_processor.cli import main
    main()
