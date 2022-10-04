*&---------------------------------------------------------------------*
*&  Include           ZACETIN_ODEV2_F01
*&---------------------------------------------------------------------*

*PERFORM get_data.



*&---------------------------------------------------------------------*
*&      Form  GET_DATA
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM get_data .


  CALL FUNCTION 'ALSM_EXCEL_TO_INTERNAL_TABLE'
    EXPORTING
      filename                = p_file
      i_begin_col             = 1
      i_begin_row             = 2
      i_end_col               = 7
      i_end_row               = 1000
    TABLES
      intern                  = gt_exceldata
    EXCEPTIONS
      inconsistent_parameters = 1
      upload_ole              = 2
      OTHERS                  = 3.
  IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
* WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ENDIF.
ENDFORM.

FORM list_data .
  CALL SCREEN 0100.
ENDFORM.

FORM prepare_alv .
  DATA:  lv_fcat TYPE  slis_tabname .
  UNASSIGN <fs_data>.

  lv_fcat = 'GS_INTEXCEL'.
  ASSIGN ('GT_INTEXCEL[]') TO <fs_data>.


  IF gcl_con IS INITIAL.
    PERFORM create_container.
    PERFORM create_fcat USING lv_fcat.
*    PERFORM set_fcat .
    PERFORM set_layout.
*    PERFORM set_dropdown.
    PERFORM display_alv.
    PERFORM set_handler_events.
  ELSE.
    PERFORM refresh_alv.
  ENDIF.
ENDFORM.                    " PREPARE_ALV


*&---------------------------------------------------------------------*
*&      Form  CREATE_CONTAINER
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM create_container.

  CREATE OBJECT gcl_con
    EXPORTING
      container_name              = 'CON'
    EXCEPTIONS
      cntl_error                  = 1
      cntl_system_error           = 2
      create_error                = 3
      lifetime_error              = 4
      lifetime_dynpro_dynpro_link = 5
      OTHERS                      = 6.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
               WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.

  CREATE OBJECT gcl_grid
    EXPORTING
      i_parent          = gcl_con
    EXCEPTIONS
      error_cntl_create = 1
      error_cntl_init   = 2
      error_cntl_link   = 3
      error_dp_create   = 4
      OTHERS            = 5.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
               WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.
  "bu method edite acık gelmesı ıcın gereklı
  CALL METHOD gcl_grid->set_ready_for_input
    EXPORTING
      i_ready_for_input = 1.

ENDFORM.                    " CREATE_CONTAINER


FORM create_fcat USING p_fcat TYPE  slis_tabname .
  CLEAR: gt_fcat,gt_fcat[],it_fieldcat[].
  REFRESH: gt_fcat[].  FREE: gt_fcat[].


  CALL FUNCTION 'REUSE_ALV_FIELDCATALOG_MERGE'
    EXPORTING
      i_program_name         = sy-repid
      i_internal_tabname     = p_fcat
      i_inclname             = 'ZACETIN_ODEV2_TOP'
    CHANGING
      ct_fieldcat            = it_fieldcat
    EXCEPTIONS
      inconsistent_interface = 1
      program_error          = 2
      OTHERS                 = 3.
  IF sy-subrc <> 0.
* Implement suitable error handling here
  ENDIF.


  CALL FUNCTION 'LVC_TRANSFER_FROM_SLIS'
    EXPORTING
      it_fieldcat_alv = it_fieldcat
    IMPORTING
      et_fieldcat_lvc = gt_fcat
    TABLES
      it_data         = gt_intexcel[]
    EXCEPTIONS
      it_data_missing = 1
      OTHERS          = 2.
  IF sy-subrc <> 0.
* Implement suitable error handling here
  ENDIF.


ENDFORM.                    " CREATE_FCAT


FORM display_alv .
  gs_vari-report = sy-repid.
  gs_vari-variant = ''.
*  gs_layo-cwidth_opt = 'X'.
*  gs_layo-zebra = 'X'.
*  gs_layo-stylefname = 'STYLE'.

  CALL METHOD gcl_grid->set_table_for_first_display
    EXPORTING
      is_variant                    = gs_vari
      i_save                        = 'A'
      is_layout                     = gs_layo
    CHANGING
      it_outtab                     = <fs_data>
      it_fieldcatalog               = gt_fcat
    EXCEPTIONS
      invalid_parameter_combination = 1
      program_error                 = 2
      too_many_lines                = 3
      OTHERS                        = 4.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
               WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.
  gcl_grid->set_ready_for_input(
         i_ready_for_input = 1
     ).
ENDFORM.                    " DISPLAY_ALV


*FORM set_fcat.
*  FIELD-SYMBOLS : <fcat> TYPE lvc_s_fcat.
**  IF gv_okcode EQ 'SIMULATE' OR gv_okcode EQ 'SAVE'.
*
*  LOOP AT gt_fcat ASSIGNING <fcat>.
*
*      CASE <fcat>-fieldname.
*        WHEN 'LINE_COLOR' or 'CELL_COLOR'.
*          <fcat>-tech = 'X'.
*        WHEN 'CARRID'.
*          <fcat>-hotspot = 'X'.
*        WHEN OTHERS.
*      ENDCASE.
*
*  ENDLOOP.
*
*ENDFORM.                    " BUILD_FCAT

FORM set_layout.
  gs_layo-stylefname = 'STYLE'.
  gs_layo-cwidth_opt = 'X'.
  gs_layo-zebra = 'X'.

  gs_layo-no_rowins = 'X'.
  gs_layo-no_rowmove = 'X'.

ENDFORM.                    " SET_LAYOUT

FORM set_handler_events .
  CREATE OBJECT gcl_evt_rec.
  SET HANDLER gcl_evt_rec->handle_user_command
          FOR gcl_grid.

  SET HANDLER gcl_evt_rec->handle_toolbar
        FOR gcl_grid.

  CALL METHOD gcl_grid->register_edit_event
    EXPORTING
      i_event_id = cl_gui_alv_grid=>mc_evt_enter.
  CALL METHOD gcl_grid->set_toolbar_interactive.




ENDFORM.                    " SET_EVENTS

FORM refresh_alv .
  CALL METHOD gcl_grid->refresh_table_display
    EXPORTING
      is_stable      = gs_stbl
      i_soft_refresh = gs_soft_ref
    EXCEPTIONS
      finished       = 1
      OTHERS         = 2.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
               WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.
ENDFORM.                    " REFRESH_ALV


FORM free .
  IF gcl_grid IS NOT INITIAL.
    CALL METHOD gcl_grid->free
      EXCEPTIONS
        cntl_error        = 1
        cntl_system_error = 2
        OTHERS            = 3.
    IF sy-subrc NE 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
      WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ELSE.
      FREE gcl_grid.
    ENDIF.
  ENDIF.
  IF gcl_con IS NOT INITIAL.
    CALL METHOD gcl_con->free
      EXCEPTIONS
        cntl_error        = 1
        cntl_system_error = 2
        OTHERS            = 3.
    IF sy-subrc NE 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
      WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ELSE.
      FREE gcl_con.
    ENDIF.
  ENDIF.
ENDFORM.                    " FREE


FORM handle_user_command USING e_ucomm TYPE sy-ucomm
                               e_sender TYPE REF TO cl_gui_alv_grid...
  DATA: lv_ucomm LIKE sy-ucomm,
        lv_subrc LIKE sy-subrc.
*
*
  CALL METHOD gcl_grid->get_selected_rows
    IMPORTING
      et_row_no = gt_row_no.
*
*
  IF lines( gt_row_no ) < 1.
    MESSAGE s001(zegt_acetin) DISPLAY LIKE 'E'.
    RETURN.
*  ELSEIF lines( lt_row_no ) > 1.
*    MESSAGE s002(zace) DISPLAY LIKE 'E'..
*    RETURN.
  ENDIF.
*
  "bu alv nın ustundekı ekledıgımız butonların dustugu method
  CASE e_ucomm.


    WHEN 'SAVE'.
      CLEAR gs_intexcel.


*      PERFORM print TABLES lt_row_no USING e_ucomm.
      LOOP AT gt_row_no INTO gs_row_no.
        READ TABLE gt_intexcel INTO gs_intexcel  INDEX gs_row_no-row_id .  "indexe göre tablodaki seçilen satırı alıyor.
        IF sy-subrc EQ 0.
          SELECT COUNT(*) FROM zacetin_excel_t
            WHERE matnr = gs_intexcel-matnr.
          IF sy-subrc EQ 0 .
            gs_message-id = 'ZEGT_ACETIN'. "se91 deki oluşturduğun mesaj classının ismi
            gs_message-type = 'E'. "success mesajının S si.
            gs_message-number = 003. "mesaj classındaki mesajın numarası
            gs_message-message_v1 = gs_intexcel-maktx. "eğer mesaja & koyduysan oraya geçecek şey.
          ELSE.
            INSERT INTO zacetin_excel_t VALUES gs_intexcel.
            gs_message-id = 'ZEGT_ACETIN'. "se91 deki oluşturduğun mesaj classının ismi
            gs_message-type = 'S'. "success mesajının S si.
            gs_message-number = 002. "mesaj classındaki mesajın numarası
            gs_message-message_v1 = gs_intexcel-maktx. "eğer mesaja & koyduysan oraya geçecek şey.
          ENDIF.


        ENDIF.
      ENDLOOP.
      "ALV BASTIKTAN SONRA WRİTE KOMUTU ÇALIŞMAZ.


      CALL FUNCTION 'POPUP_DISPLAY_MESSAGE'
        EXPORTING
*         TITLE =
          msgid = gs_message-id
          msgty = gs_message-type
          msgno = gs_message-number
          msgv1 = gs_message-message_v1
*    *           MSGV2 =
*    *           MSGV3 =
*    *           MSGV4 =
*    *           START_COLUMN       = 5
*    *           START_ROW          = 5
        .



  ENDCASE.
  PERFORM refresh_alv.


**   perform add_messa
**  PERFORM show_message.
*  CALL FUNCTION 'OXT_MESSAGE_TO_POPUP'
*    EXPORTING
*      it_message =
**   IMPORTING
**     EV_CONTINUE       =
**   EXCEPTIONS
**     BAL_ERROR  = 1
**     OTHERS     = 2
*    .
*  IF sy-subrc <> 0.
** Implement suitable error handling here
*  ENDIF.

ENDFORM.                    " HANDLE_USER_COMMAND



*&---------------------------------------------------------------------*
*&      Form  MOVE_DATA
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM move_data .

  DATA: lv_index TYPE i.
  lv_index = 1.

  LOOP AT gt_exceldata INTO gs_exceldata.
    IF gs_exceldata-row EQ lv_index.
      CASE gs_exceldata-col.
        WHEN 1.
          gs_intexcel-matnr = gs_exceldata-value.
        WHEN 2.
          gs_intexcel-maktx = gs_exceldata-value.
        WHEN 3.
          gs_intexcel-menge = gs_exceldata-value.
        WHEN 4.

          gs_intexcel-meins = gs_exceldata-value.

          CALL FUNCTION 'CONVERSION_EXIT_CUNIT_INPUT'
            EXPORTING
              input          = gs_intexcel-meins
              language       = sy-langu
            IMPORTING
              output         = gs_intexcel-meins
            EXCEPTIONS
              unit_not_found = 1
              OTHERS         = 2.
          IF sy-subrc <> 0.
* Implement suitable error handling here
          ENDIF.

        WHEN 5.
          gs_intexcel-dmbtr = gs_exceldata-value.
        WHEN 6.
          gs_intexcel-waers = gs_exceldata-value.
        WHEN 7.
          gs_intexcel-tarih = gs_exceldata-value.
          lv_index = lv_index + 1.
          APPEND gs_intexcel TO gt_intexcel.
        WHEN OTHERS.
      ENDCASE.

    ENDIF.


  ENDLOOP.

ENDFORM.
*&---------------------------------------------------------------------*
*&      Form  F4_FILE
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM f4_file .
  MOVE '(.xlsx)|.xlsx|' TO  file_filter.


  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
      window_title            = 'Dosyayı Seçiniz !'
      file_filter             = 'EXCEL FILES (*.XLSX)|*.XLSX|'
    CHANGING
      file_table              = file_table
      rc                      = rc
    EXCEPTIONS
      file_open_dialog_failed = 1
      cntl_error              = 2
      error_no_gui            = 3
      not_supported_by_gui    = 4
      OTHERS                  = 5.

  IF sy-subrc <> 0.
    EXIT.
  ENDIF.

  IF file_table[] IS INITIAL.
    EXIT.
  ENDIF.

  READ TABLE file_table INTO fl INDEX 1.
  IF sy-subrc NE 0.
    READ TABLE file_table INTO fl INDEX 2.
  ENDIF.

  p_file = fl-filename.
  DATA : ln TYPE i.

  ln = strlen( p_file ) - 4.

  IF p_file+ln(4) NE 'xlsx' .
    MESSAGE 'xlsx  uzantılı dosya seçiniz' TYPE 'I'.
    EXIT.
  ENDIF.

  DATA: dyfield  TYPE  dynpread,
        dyfields TYPE STANDARD TABLE OF dynpread.

  dyfield-fieldname = 'P_FILE'.
  dyfield-fieldvalue = p_file.
  APPEND dyfield TO dyfields.
  CLEAR dyfield.
  dyfield-fieldname = 'P_FILE'.
  dyfield-fieldvalue = p_file .
  APPEND dyfield TO dyfields.

  CALL FUNCTION 'DYNP_VALUES_UPDATE'
    EXPORTING
      dyname               = sy-cprog
      dynumb               = sy-dynnr
    TABLES
      dynpfields           = dyfields
    EXCEPTIONS
      invalid_abapworkarea = 1
      invalid_dynprofield  = 2
      invalid_dynproname   = 3
      invalid_dynpronummer = 4
      invalid_request      = 5
      no_fielddescription  = 6
      undefind_error       = 7
      OTHERS               = 8.

  IF sy-subrc <> 0.
* Implement suitable error handling here
  ENDIF.

ENDFORM.


FORM handle_toolbar  USING e_object TYPE REF TO cl_alv_event_toolbar_set
                           e_sender TYPE REF TO cl_gui_alv_grid.

  DATA  : lt_toolbar TYPE ttb_button.
  DATA  : ls_toolbar TYPE stb_button.
  CASE e_sender.
    WHEN gcl_grid.

      ls_toolbar-function  = 'SAVE'."'Kaydet
      ls_toolbar-quickinfo = TEXT-002."YKaydet
      ls_toolbar-butn_type = '0'.
      ls_toolbar-text      = TEXT-002."Kaydet
      ls_toolbar-icon      = icon_system_save.
      INSERT ls_toolbar INTO TABLE lt_toolbar .



  ENDCASE.

  INSERT LINES OF lt_toolbar INTO TABLE  e_object->mt_toolbar[].
*  e_object->mt_toolbar[] = lt_toolbar[].
ENDFORM.                    " HANDLE_TOOLBAR
*&---------------------------------------------------------------------*
*&      Form  DOWN_EXCEL
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM down_excel .


    DATA: excel    TYPE ole2_object,
        workbook TYPE ole2_object,
        sheet    TYPE ole2_object,
        cell     TYPE ole2_object,
        row      TYPE ole2_object.

  CREATE OBJECT excel 'EXCEL.APPLICATION'.

*  Create an Excel workbook Object.
  CALL METHOD OF excel 'WORKBOOKS' = workbook .

*  Put Excel in background
  SET PROPERTY OF excel 'VISIBLE' = 0 .

*  Put Excel in front.
  SET PROPERTY OF excel 'VISIBLE' = 1 .

  CALL METHOD OF workbook 'add'.

  CALL METHOD OF excel 'Worksheets' = sheet
    EXPORTING
    #1 = 1.

  CALL METHOD OF sheet 'Activate'.

  " Excel Raw Header Line
*  CALL METHOD OF excel 'RANGE' = cell
*    EXPORTING
*    #1 = 'A1'
*    #2 = 'J1'.
*  CALL METHOD OF cell 'MERGE'.
*  SET PROPERTY OF cell 'VALUE' = text-h01.
*
*  CALL METHOD OF excel 'RANGE' = cell
*    EXPORTING
*    #1 = 'K1'.
*  SET PROPERTY OF cell 'VALUE' = text-h02.
*
*  CALL METHOD OF excel 'RANGE' = cell
*    EXPORTING
*    #1 = 'L1'
*    #2 = 'M1'.
*  CALL METHOD OF cell 'MERGE'.
*  SET PROPERTY OF cell 'VALUE' = text-h03.
*
*  CALL METHOD OF excel 'RANGE' = cell
*    EXPORTING
*    #1 = 'N1'
*    #2 = 'Q1'.
*  CALL METHOD OF cell 'MERGE'.
*  SET PROPERTY OF cell 'VALUE' = text-h04.
*
*  CALL METHOD OF excel 'RANGE' = cell
*    EXPORTING
*    #1 = 'R1'.
*  SET PROPERTY OF cell 'VALUE' = text-h05.
*
*  CALL METHOD OF excel 'RANGE' = cell
*    EXPORTING
*    #1 = 'S1'.
*  SET PROPERTY OF cell 'VALUE' = text-h06.

  CALL METHOD OF excel 'RANGE' = cell
    EXPORTING
    #1 = 'A1'.
  SET PROPERTY OF cell 'VALUE' = 'MALZEME'.

  CALL METHOD OF excel 'RANGE' = cell
    EXPORTING
    #1 = 'B1'.
  SET PROPERTY OF cell 'VALUE' = 'MALZEME MİKTARI'.

  CALL METHOD OF excel 'RANGE' = cell
    EXPORTING
    #1 = 'C1'.
  SET PROPERTY OF cell 'VALUE' = 'TALEP MİKTARI'.

  CALL METHOD OF excel 'RANGE' = cell
    EXPORTING
    #1 = 'D1'.
  SET PROPERTY OF cell 'VALUE' = 'BİRİM'.

  CALL METHOD OF excel 'RANGE' = cell
    EXPORTING
    #1 = 'E1'.
  SET PROPERTY OF cell 'VALUE' = 'BİRİM FİYATI'.

  CALL METHOD OF excel 'RANGE' = cell
    EXPORTING
    #1 = 'F1'.
  SET PROPERTY OF cell 'VALUE' = 'PARA BİRİMİ'.

  CALL METHOD OF excel 'RANGE' = cell
    EXPORTING
    #1 = 'G1'.
  SET PROPERTY OF cell 'VALUE' = 'TARİH'.


*  " Excel Sample Row 1
*  ls_excel_format-matnr = '0000000001'.
*  ls_excel_format-werks = '1000'.
*  ls_excel_format-cost  = '197549'.
*
*  append ls_excel_format to lt_excel_format.
*
*  " Append Excel Sample Data to Internal Table.
*  loop at lt_excel_format.
*    call method of excel 'ROWS' = row
*      exporting
*        #1 = '2'.
*    call method of row 'INSERT' no flush.
*    call method of excel 'RANGE' = cell
*      exporting
*        #1 = 'A2'.
*    set property of cell 'VALUE' = lt_excel_format-matnr no flush.
*    call method of excel 'RANGE' = cell
*      exporting
*        #1 = 'B2'.
*    set property of cell 'VALUE' = lt_excel_format-werks no flush.
*    call method of excel 'RANGE' = cell
*      exporting
*        #1 = 'C2'.
*    set property of cell 'VALUE' = lt_excel_format-cost no flush.
*
*  endloop.

  CALL METHOD OF excel 'SAVE'.
  CALL METHOD OF excel 'QUIT'.

*  Free all objects
  FREE OBJECT cell.
  FREE OBJECT workbook.
  FREE OBJECT excel.
  excel-handle = -1.
  FREE OBJECT row.
ENDFORM.
