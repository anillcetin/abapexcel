*&---------------------------------------------------------------------*
*& Report ZACETIN_ODEV2
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zacetin_odev2.

TABLES sscrfields.

INCLUDE zacetin_odev2_top.
INCLUDE zacetin_odev2_cls.
INCLUDE zacetin_odev2_f01.
INCLUDE zacetin_odev2_pbo.
INCLUDE zacetin_odev2_pai.


SELECTION-SCREEN FUNCTION KEY 1.
*
INITIALIZATION.
CONCATENATE icon_print
            'Excel Şablonu İndir'
            INTO sscrfields-functxt_01 SEPARATED BY space.


*buton
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM f4_file.

AT SELECTION-SCREEN.
  IF sscrfields-ucomm = 'FC01'.
    perform down_excel.

  ENDIF.



START-OF-SELECTION.
  PERFORM get_data.
  PERFORM : move_data.

END-OF-SELECTION.
  PERFORM list_data.
