***&---------------------------------------------------------------------*
***&  Include           ZACETIN_ODEV2_CLS
***&---------------------------------------------------------------------*
**
CLASS lcl_alv DEFINITION .
  PUBLIC SECTION.
    METHODS:

      handle_user_command FOR EVENT user_command   OF cl_gui_alv_grid
        IMPORTING e_ucomm
                  sender,

      handle_toolbar      FOR EVENT toolbar        OF cl_gui_alv_grid
        IMPORTING e_object
                  sender.

ENDCLASS.
**
CLASS lcl_alv IMPLEMENTATION.
  METHOD handle_user_command.
    PERFORM handle_user_command
      USING e_ucomm
            sender.
  ENDMETHOD.

  METHOD handle_toolbar.
    PERFORM handle_toolbar
      USING e_object
            sender.
  ENDMETHOD.                    "handle_toolbar


ENDCLASS.
**
**  METHOD get_data.
**    PERFORM get_data.
**    PERFORM : move_data.
**  ENDMETHOD.
**  METHOD list_data .
**    PERFORM list_data.
**  ENDMETHOD.
**
**ENDCLASS.
