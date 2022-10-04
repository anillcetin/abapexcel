*&---------------------------------------------------------------------*
*&  Include           ZACETIN_ODEV2_TOP
*&---------------------------------------------------------------------*
 TABLES: mara,makt.

 TYPE-POOLS: slis, icon.
 DATA: gv_okcode TYPE sy-ucomm.





 FIELD-SYMBOLS: <fs_data> TYPE ANY TABLE.
 "STRUCTUREDAN WİTH HEADER LİNE İLE AL.
 DATA:gs_row_no TYPE lvc_s_roid,
      gt_row_no TYPE lvc_t_roid.
 "STRUCTUREDAN WİTH HEADER LİNE İLE AL.

*>*ALV Tanımlamaları
CLASS : lcl_alv     DEFINITION DEFERRED.
DATA  : gcl_evt_rec TYPE REF TO lcl_alv.
DATA  : gcl_alv     TYPE REF TO lcl_alv,
        gcl_grid    TYPE REF TO cl_gui_alv_grid,
        gcl_con     TYPE REF TO cl_gui_custom_container,
        gcl_docking TYPE REF TO cl_gui_docking_container.
DATA  : it_fieldcat TYPE slis_t_fieldcat_alv.
DATA  : gt_fcat     TYPE lvc_t_fcat.
DATA  : gs_fcat     TYPE lvc_s_fcat.
DATA  : gs_layo     TYPE lvc_s_layo.
DATA  : gs_vari     TYPE disvariant.
DATA  : gs_stbl     TYPE lvc_s_stbl.
DATA  : gs_soft_ref TYPE char1 VALUE 'X'.
DATA  : gv_declaration_no TYPE char20.
DATA  : gt_drop TYPE lvc_t_drop.



 DATA: fm_name         TYPE rs38l_fnam,
       fp_docparams    TYPE sfpdocparams,
       fp_outputparams TYPE sfpoutputparams,
       formname        TYPE fpname,
       fp_formoutput   TYPE fpformoutput.

***popup message,
DATA: gs_message  LIKE bapiret2.
DATA: gt_message  LIKE TABLE OF gs_message.


******************************************************

 DATA: gs_intexcel LIKE zacetin_excel_t,
       gt_intexcel LIKE TABLE OF gs_intexcel.




 DATA: file_table  TYPE filetable,
       fl          LIKE LINE OF file_table,
       rc          TYPE i,
       file_filter TYPE    string.

 DATA: gt_exceldata LIKE TABLE OF alsmex_tabline,
       gs_exceldata LIKE  alsmex_tabline.



 SELECTION-SCREEN BEGIN OF BLOCK blk0.
 PARAMETERS p_file TYPE rlgrap-filename.
 SELECTION-SCREEN BEGIN OF LINE.
*    SELECTION-SCREEN PUSHBUTTON  2(30) TEXT-005 USER-COMMAND bt1.
 SELECTION-SCREEN END OF LINE.
 SELECTION-SCREEN END OF BLOCK blk0.
