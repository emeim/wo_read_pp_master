*&---------------------------------------------------------------------*
*& Report ZRPP_IDP_0478
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zrpp_idp_0478a.

*&---------------------------------------------------------------------*
*& Report ZRPP_IDP_0478
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*

TABLES: sscrfields,caufv.
DATA: lt_file        TYPE filetable,
      ls_file        LIKE LINE OF lt_file,
      lv_rc          TYPE i,
      lv_user_action TYPE i,
      lv_file_filter TYPE string.

DATA: BEGIN OF t_log OCCURS 0,
        aufnr        LIKE caufv-aufnr,
        flag(1)      TYPE c,
        message(100) TYPE c,
      END OF t_log.
TYPES: BEGIN OF t_input,
         werks LIKE caufv-werks,
         aufnr LIKE caufv-aufnr,

       END OF t_input.
DATA: input TYPE STANDARD TABLE OF t_input WITH EMPTY KEY."  WITH HEADER LINE.
DATA: input1 TYPE STANDARD TABLE OF t_input WITH HEADER LINE.

SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE txt01.
  PARAMETERS: p_werks LIKE aufk-werks OBLIGATORY DEFAULT 'AEY1',
              p_file  LIKE rlgrap-filename DEFAULT 'D:\zdppp478.xlsx' OBLIGATORY,
              runmode TYPE c DEFAULT 'E'.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN FUNCTION KEY 1.

INCLUDE zpp_bdc.

INITIALIZATION.
  txt01 = '选择画面'.
  sscrfields-functxt_01 = '下载批导模板'.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM open_file.

AT SELECTION-SCREEN.
  PERFORM download_template.


START-OF-SELECTION.
  PERFORM check_auth.
*  PERFORM upload_file.
  PERFORM upload_file1.
  PERFORM check_data.
  PERFORM process_data.
  PERFORM write_log.

*&---------------------------------------------------------------------*
*& Form open_file
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM open_file .
  CONCATENATE 'Excel (*.xls;*.xlsx)|*.xls;*.xlsx'
                  '|'
                  'All Files (*.*)|*.*'
                  INTO lv_file_filter.
  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
*     window_title            =
*     default_extension       =
*     default_filename        =
      file_filter             = lv_file_filter
*     with_encoding           =
      initial_directory       = 'D:\'
*     multiselection          =
    CHANGING
      file_table              = lt_file
      rc                      = lv_rc
      user_action             = lv_user_action
*     file_encoding           =
    EXCEPTIONS
      file_open_dialog_failed = 1
      cntl_error              = 2
      error_no_gui            = 3
      not_supported_by_gui    = 4
      OTHERS                  = 5.
  IF sy-subrc <> 0.
    MESSAGE 'File Open failed' TYPE 'E' RAISING error.      " File Open failed
  ENDIF.

  IF lv_user_action EQ cl_gui_frontend_services=>action_cancel.

    RETURN.
  ENDIF.

  READ TABLE lt_file INTO ls_file INDEX 1.
  IF sy-subrc = 0.
    p_file  = ls_file-filename.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form upload_file
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM upload_file .
  DATA : lv_filename      TYPE string,
         lt_records       TYPE solix_tab,
         lv_headerxstring TYPE xstring,
         lv_filelength    TYPE i.

  FIELD-SYMBOLS : <gt_data>       TYPE STANDARD TABLE .
  FIELD-SYMBOLS : <ls_data>  TYPE any,
                  <lv_field> TYPE any.
  lv_filename = p_file.
  CALL METHOD cl_gui_frontend_services=>gui_upload
    EXPORTING
      filename                = lv_filename
      filetype                = 'BIN'
*     has_field_separator     = 'X'
*     header_length           = 0
*     read_by_line            = 'X'
*     dat_mode                = SPACE
    IMPORTING
      filelength              = lv_filelength
      header                  = lv_headerxstring
    CHANGING
      data_tab                = lt_records
    EXCEPTIONS
      file_open_error         = 1
      file_read_error         = 2
      no_batch                = 3
      gui_refuse_filetransfer = 4
      invalid_type            = 5
      no_authority            = 6
      unknown_error           = 7
      bad_data_format         = 8
      header_not_allowed      = 9
      separator_not_allowed   = 10
      header_too_long         = 11
      unknown_dp_error        = 12
      access_denied           = 13
      dp_out_of_memory        = 14
      disk_full               = 15
      dp_timeout              = 16
      not_supported_by_gui    = 17
      error_no_gui            = 18
      OTHERS                  = 19.
  CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
    EXPORTING
      input_length = lv_filelength
    IMPORTING
      buffer       = lv_headerxstring
    TABLES
      binary_tab   = lt_records
    EXCEPTIONS
      failed       = 1
      OTHERS       = 2.
  IF sy-subrc <> 0.
    "Implement suitable error handling here
  ENDIF.
  DATA : lo_excel_ref TYPE REF TO cl_fdt_xl_spreadsheet .

  TRY .
      lo_excel_ref = NEW cl_fdt_xl_spreadsheet(
                              document_name = lv_filename
                              xdocument     = lv_headerxstring ) .
    CATCH cx_fdt_excel_core.
      "Implement suitable error handling here
  ENDTRY .
  lo_excel_ref->if_fdt_doc_spreadsheet~get_worksheet_names(
   IMPORTING
     worksheet_names = DATA(lt_worksheets) ).

  IF NOT lt_worksheets IS INITIAL.
    READ TABLE lt_worksheets INTO DATA(lv_woksheetname) INDEX 1.

    DATA(lo_data_ref) = lo_excel_ref->if_fdt_doc_spreadsheet~get_itab_from_worksheet(
                                             lv_woksheetname ).
    "now you have excel work sheet data in dyanmic internal table
    ASSIGN lo_data_ref->* TO <gt_data>.
  ENDIF.

  DATA : lv_numberofcolumns   TYPE i,
         lv_date_string       TYPE string,
         lv_target_date_field TYPE datum,
         lt_dataset           TYPE TABLE OF t_input,
         ls_dataset           TYPE t_input.
  lv_numberofcolumns = 2.
  LOOP AT <gt_data> ASSIGNING <ls_data> FROM 2.
    CLEAR ls_dataset.
    DO lv_numberofcolumns TIMES.
      ASSIGN COMPONENT sy-index OF STRUCTURE <ls_data> TO <lv_field> .
      IF sy-subrc = 0.
        CASE sy-index.
          WHEN 1.
            ls_dataset-werks = <lv_field>.
          WHEN 2.
            ls_dataset-aufnr = <lv_field>.

        ENDCASE.
      ENDIF.
    ENDDO.
    APPEND ls_dataset TO lt_dataset.
  ENDLOOP.
  input[] = lt_dataset[].
  DELETE input WHERE aufnr EQ ''.
  REFRESH lt_dataset.
  IF input[] IS INITIAL.
    MESSAGE i003(zmm001).
    LEAVE LIST-PROCESSING.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form download_template
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM download_template .

  DATA: lv_file_name         TYPE string,
        lv_path              TYPE string,
        lv_fullpath          TYPE string,
        lv_file_filter       TYPE string,
        lv_user_action       TYPE i,
        lv_default_file_name TYPE string.

  DATA: lv_file TYPE rlgrap-filename.
  DATA: lv_wwwdatatab TYPE wwwdatatab VALUE 'ZRPP_IDP_0478'.

  CONCATENATE 'Excel (*.xls;*.xlsx)|*.xls;*.xlsx'
                '|'
                'All Files (*.*)|*.*'
                INTO lv_file_filter.
  lv_default_file_name = '批量反结案工单的模板.xlsx'.
  CASE sscrfields-ucomm.
    WHEN 'FC01'.
      CALL METHOD cl_gui_frontend_services=>file_save_dialog
        EXPORTING
          window_title      = '批量反结案工单的的模板'
*         default_extension =
          default_file_name = lv_default_file_name
*         with_encoding     =
          file_filter       = lv_file_filter
*         initial_directory =
*         prompt_on_overwrite  = 'X'
        CHANGING
          filename          = lv_file_name
          path              = lv_path
          fullpath          = lv_fullpath
          user_action       = lv_user_action
*         file_encoding     =
*    EXCEPTIONS
*         cntl_error        = 1
*         error_no_gui      = 2
*         not_supported_by_gui = 3
*         others            = 4
        .
      IF sy-subrc <> 0.
*   Implement suitable error handling here
      ENDIF.
      IF lv_user_action = cl_gui_frontend_services=>action_cancel.

      ELSE.
        lv_file = lv_fullpath.
        SELECT SINGLE * INTO CORRESPONDING FIELDS OF lv_wwwdatatab FROM wwwdata WHERE objid = 'ZRPP_IDP_0478' .

        "下载模板到指定路径
        CALL FUNCTION 'DOWNLOAD_WEB_OBJECT'
          EXPORTING
            key         = lv_wwwdatatab
            destination = lv_file.
*           IMPORTING
*           RC          =
*           CHANGING
*           TEMP        =
      ENDIF.

    WHEN OTHERS.
  ENDCASE.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form check_auth
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM check_auth .
  AUTHORITY-CHECK OBJECT 'C_AFKO_AWA'
          ID 'ACTVT' FIELD '02'
          ID 'WERKS' FIELD p_werks.
  IF sy-subrc NE 0.
    MESSAGE i000(zpp) WITH '无工厂生产订单修改权限'.
    STOP.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form check_data
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM check_data .
  DATA: lt_order_r TYPE TABLE OF bapi_pp_orderrange,
        ls_order_r LIKE LINE OF  lt_order_r,
        lt_ord_hdr TYPE TABLE OF bapi_order_header1,
        ls_ord_hdr LIKE LINE OF  lt_ord_hdr.
  LOOP AT input ASSIGNING FIELD-SYMBOL(<fs_input>).
    IF <fs_input>-werks NE p_werks.
      MESSAGE i000(zpp) WITH <fs_input>-werks ' 工厂不一致,请确认!'.
      STOP.
    ENDIF.
    CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
      EXPORTING
        input  = <fs_input>-aufnr
      IMPORTING
        output = <fs_input>-aufnr.
    SELECT SINGLE * FROM caufv WHERE aufnr = <fs_input>-aufnr.
    IF sy-subrc NE 0.
      MESSAGE i000(zpp) WITH <fs_input>-aufnr ' 工单号码不存在,请确认!'.
      STOP.
    ENDIF.
  ENDLOOP.

  lt_order_r = VALUE #( FOR ls_input IN input
                         ( sign = 'I' option = 'EQ' low = ls_input-aufnr ) ).

  CALL FUNCTION 'BAPI_PRODORD_GET_LIST'
    TABLES
      order_number_range = lt_order_r
      order_header       = lt_ord_hdr.

  LOOP AT lt_ord_hdr INTO ls_ord_hdr.
    IF ls_ord_hdr-system_status NS 'TECO'.
*      MESSAGE i000(zpp) WITH ls_ord_hdr-order_number ' 未结案,请确认'.
*      STOP.
      DELETE input WHERE aufnr = ls_ord_hdr-order_number.
    ENDIF.
  ENDLOOP.
  IF input[] IS INITIAL.
    MESSAGE i000(zpp) WITH '无满足条件的工单,请确认!'.
    STOP.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form process_data
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM process_data .
  DATA: lt_bapi_return TYPE TABLE OF bapiret2 WITH HEADER LINE.

  SELECT a~aufnr ,gamng, gmein ,gltrp, gstrp ,terkz ,fhori FROM caufv AS a
    INNER JOIN @input AS b ON a~werks = b~werks AND a~aufnr = b~aufnr
  INTO TABLE @DATA(lt_aufnr).

  LOOP AT lt_aufnr ASSIGNING FIELD-SYMBOL(<fs_aufnr>).
    REFRESH: bdcdata.
    PERFORM bdc_dynpro      USING 'SAPLCOKO1' '0110'.
    PERFORM bdc_field       USING 'BDC_CURSOR'
                                  'CAUFVD-AUFNR'.
    PERFORM bdc_field       USING 'BDC_OKCODE'
                                  '=ENTK'.
    PERFORM bdc_field       USING 'CAUFVD-AUFNR'
                                  <fs_aufnr>-aufnr.
    PERFORM bdc_field       USING 'R62CLORD-FLG_OVIEW'
                                  'X'.
    PERFORM bdc_dynpro      USING 'SAPLCOKO1' '0115'.
    PERFORM bdc_field       USING 'BDC_OKCODE'
                                  '=TABR'.
    PERFORM bdc_field       USING 'BDC_CURSOR'
                                  'CAUFVD-GAMNG'.
    PERFORM bdc_dynpro      USING 'SAPLCOKO1' '0115'.
    PERFORM bdc_field       USING 'BDC_OKCODE'
                                  '=BU'.
    PERFORM bdc_field       USING 'BDC_CURSOR'
                                  'CAUFVD-GAMNG'.
    PERFORM bdc_field       USING 'CAUFVD-GAMNG'
                                   <fs_aufnr>-gamng.
    PERFORM bdc_field       USING 'CAUFVD-GLTRP'
                                  <fs_aufnr>-gltrp.
    PERFORM bdc_field       USING 'CAUFVD-GSTRP'
                                  <fs_aufnr>-gstrp.
    PERFORM bdc_field       USING 'CAUFVD-TERKZ'
                                   <fs_aufnr>-terkz.
    PERFORM bdc_field       USING 'CAUFVD-FHORI'
                                   <fs_aufnr>-fhori.

    REFRESH: messtab,lt_bapi_return.
    CALL TRANSACTION 'CO02' USING bdcdata
                         UPDATE  'S'
                         MODE runmode
                         MESSAGES INTO messtab.
    CALL FUNCTION 'CONVERT_BDCMSGCOLL_TO_BAPIRET2'
      TABLES
        imt_bdcmsgcoll = messtab
        ext_return     = lt_bapi_return.

    CLEAR t_log.
    t_log-aufnr = <fs_aufnr>-aufnr.

    LOOP AT lt_bapi_return INTO DATA(ls_bapi_return) WHERE ( type = 'E' OR type = 'A' ).
      t_log-flag = 'E'.
      t_log-message = '反转失败!' && ls_bapi_return-message.

    ENDLOOP.
    IF t_log-flag = ''.
      t_log-flag = 'S'.
      READ TABLE lt_bapi_return INTO ls_bapi_return INDEX 1.
      t_log-message = '反转成功!' && ls_bapi_return-message.
    ENDIF.
    APPEND t_log.
  ENDLOOP.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form write_log
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM write_log .
  LOOP AT t_log WHERE flag = 'E'.
    WRITE: / t_log-aufnr , t_log-message COLOR 6.
  ENDLOOP.
  LOOP AT t_log WHERE flag = 'S'.
    WRITE: / t_log-aufnr , t_log-message COLOR 5.
  ENDLOOP.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form upload_file1
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM upload_file1 .
  DATA intern TYPE STANDARD TABLE OF zalsmex_tabline WITH HEADER LINE.
  FIELD-SYMBOLS:<fs>.
  CALL FUNCTION 'ZALSM_EXCEL_TO_INTERNAL_TABLE'
    EXPORTING
      filename                = p_file
      i_begin_col             = 1
      i_begin_row             = 2
      i_end_col               = 255
      i_end_row               = 65336
    TABLES
      intern                  = intern
    EXCEPTIONS
      inconsistent_parameters = 1
      upload_ole              = 2
      OTHERS                  = 3.

  IF sy-subrc <> 0.
    MESSAGE i368(00) WITH '上传失败,请调整模板'.
    STOP.
  ENDIF.
  IF intern[] IS INITIAL.
    MESSAGE i208(00) WITH 'No Data Upload'.
    STOP.
  ENDIF.
  LOOP AT intern.

    ASSIGN COMPONENT intern-col OF STRUCTURE input1 TO <fs>.
    MOVE intern-value TO <fs>.

    AT END OF zrow.
      APPEND input1.CLEAR input1.
    ENDAT.
  ENDLOOP.
  input[] = input1[].
ENDFORM.
