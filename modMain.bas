Attribute VB_Name = "modMain"
Option Explicit


Global gbKeys(256) As Boolean       'Indicates which keys are currently pressed

Global gflXRot As GLfloat           'X rotation
Global gflYRot As GLfloat           'Y rotation
Global gflXSpeed As GLfloat         'X rotation speed
Global gflYSpeed As GLfloat         'Y rotation speed
Global gflZ As GLfloat              'Z position
