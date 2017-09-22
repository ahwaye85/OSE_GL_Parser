
from cx_Freeze import setup, Executable

setup(name="OSE GL Parser",
      Version="1.0",
      description = "Parse OSE files",
      executables = [Executable("OSE_GL_Parser.py")])





