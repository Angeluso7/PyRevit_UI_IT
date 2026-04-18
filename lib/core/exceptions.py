# -*- coding: utf-8 -*-
class PyRevitITError(Exception):
    pass

class EnvironmentError(PyRevitITError):
    pass

class ExportError(PyRevitITError):
    pass

class ValidationError(PyRevitITError):
    pass
