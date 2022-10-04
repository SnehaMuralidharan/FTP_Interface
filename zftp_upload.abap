*&---------------------------------------------------------------------*
*& Report  ZPP_ALL_UPLOAD_JOB_TICKET
*&    Title       : Program to create Multiple Job Tickets based on the excel
*&    Author      : Arul Kumar / Velraj T (CLSS)
*&    Created On  : 12-July-2019
*&    Functional
*&    Consultant  : Andrea Mietta (Tristone)
*&    Application : PP
*&    Description : This new Program will create mutiple job tickets in the TCode ZJTKT1
*&                  based on the uploaded excel file
*&    Requests    : 1.D02K938334  CLSS::PP::RAK::BN10000158::CR17251::Job ticket Interface V11
*&--------------------------------------------------------------------------------------------*
*&----------------------------------    Change History   -------------------------------------*
*&--------------------------------------------------------------------------------------------*
*&    S.No  | Changed By        |   Changed on    | Request(s)
*&      1   | xxxxxxxxxxxxxxxxx |   XX-XX-XXXX    |
*&    Reason      :
*&--------------------------------------------------------------------------------------------*

REPORT  zpp_all_upload_job_ticket MESSAGE-ID zpp_job_ticket.
TYPE-POOLS: truxs.
*------------------------------- Begin of global declarations --------------------------------&
CLASS: lcl_selscr DEFINITION DEFERRED,
       lcl_ftp DEFINITION DEFERRED,
       lcl_data DEFINITION DEFERRED,
       lcl_alv DEFINITION DEFERRED.

DATA: go_selscr TYPE REF TO lcl_selscr,
      go_ftp    TYPE REF TO lcl_ftp,
      go_data   TYPE REF TO lcl_data,
      go_alv    TYPE REF TO lcl_alv.

SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE text-003.
PARAMETERS: p_upload RADIOBUTTON GROUP rb1 USER-COMMAND rad,
            p_downld RADIOBUTTON GROUP rb1 DEFAULT 'X',
            p_backg  RADIOBUTTON GROUP rb1.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE text-001.
PARAMETERS : p_file TYPE  rlgrap-filename MODIF ID fl.
SELECTION-SCREEN BEGIN OF LINE.
SELECTION-SCREEN COMMENT 1(74) text-002 MODIF ID fl.
SELECTION-SCREEN END OF LINE.
SELECTION-SCREEN END OF BLOCK b2.

SELECTION-SCREEN BEGIN OF BLOCK b3 WITH FRAME TITLE text-004.
PARAMETERS : p_dfile TYPE  rlgrap-filename MODIF ID dl.
SELECTION-SCREEN END OF BLOCK b3.
*--------------------------------- End of global declarations --------------------------------&

*-------------------------------- Begin of Class definitions ---------------------------------&
*----------------------------------------------------------------------*
*       CLASS lcl_selscr DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_selscr DEFINITION INHERITING FROM zcl_selscreen FINAL.
  PUBLIC SECTION.
    CLASS-METHODS: get_instance RETURNING value(ro_instance) TYPE REF TO lcl_selscr.
    METHODS: read_excel_file,
             read_file_path.

    DATA: gt_excel_final TYPE TABLE OF zst_all_job_ticket.


  PRIVATE SECTION.
    CLASS-DATA: lo_selection TYPE REF TO lcl_selscr.
ENDCLASS.                    "lcl_selscr DEFINITION

*----------------------------------------------------------------------*
*       CLASS lcl_data DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_data DEFINITION FINAL.
  PUBLIC SECTION.
    CLASS-METHODS: get_instance RETURNING value(ro_instance) TYPE REF TO lcl_data.
    TYPES: BEGIN OF ty_return,
           s_no TYPE i,
           type TYPE bapi_mtype,
           id TYPE symsgid,
           number TYPE symsgno,
           message TYPE bapi_msg,
           log_no TYPE balognr,
           log_msg_no TYPE balmnr,
           message_v1 TYPE symsgv,
           message_v2 TYPE symsgv,
           message_v3 TYPE symsgv,
           message_v4 TYPE symsgv.
    TYPES:  END OF ty_return.
    DATA: gt_return	TYPE TABLE OF ty_return,
          gs_final TYPE zst_all_job_ticket,
          lt_return TYPE TABLE OF bapiret2,
          ls_return1  TYPE string,                          "#EC NEEDED
          ls_return TYPE bapiret2,
          lv_count TYPE i,
          gs_return	TYPE ty_return.

    METHODS: process_data.

  PRIVATE SECTION.

    CLASS-DATA: lo_data TYPE REF TO lcl_data.
ENDCLASS.                    "lcl_data DEFINITION

*----------------------------------------------------------------------*
*       CLASS lcl_alv DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_alv DEFINITION INHERITING FROM zcl_salv FINAL.
  PUBLIC SECTION.
    CLASS-METHODS: get_instance RETURNING value(ro_instance) TYPE REF TO lcl_alv.
    METHODS: display_alv.

  PROTECTED SECTION.
    DATA: lo_salv TYPE REF TO cl_salv_table.

  PRIVATE SECTION.
    CLASS-DATA: lo_alv TYPE REF TO lcl_alv.
ENDCLASS.                    "lcl_alv DEFINITION

*----------------------------------------------------------------------*
*       CLASS lcl_ftp DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_ftp DEFINITION FINAL.
  PUBLIC SECTION.
    CLASS-METHODS: get_instance RETURNING value(ro_instance) TYPE REF TO lcl_ftp.
    METHODS: connect_ftp,
             ftp_cmd IMPORTING gv_cmd TYPE char80,
             ftp_disconnect,
             rfc_connect,
             progress_indicator IMPORTING text TYPE char80.
    TYPES: BEGIN OF ty_result,
           line(100) TYPE c,
           END OF ty_result.

    DATA: gs_record TYPE zst_all_job_ticket,
          lv_cmd(80) TYPE c,
          result TYPE TABLE OF ty_result,
          lt_result TYPE TABLE OF ty_result,
          ls_result TYPE ty_result,
          lv_ftp_handle    TYPE i,
          i_rfc_destination  TYPE rfcdes-rfcdest.

  PRIVATE SECTION.
    CLASS-DATA: lo_ftp TYPE REF TO lcl_ftp.
ENDCLASS.                    "lcl_alv DEFINITION
*--------------------------------- End of Class definitions ----------------------------------&

*------------------------------ begin of class implementations -------------------------------&
CLASS lcl_selscr IMPLEMENTATION.
  METHOD read_excel_file.
    DATA: "gd_file TYPE rlgrap-filename,
          lt_excel TYPE STANDARD TABLE OF alsmex_tabline ,
          ls_excel TYPE alsmex_tabline,
          lt_filetable TYPE STANDARD TABLE OF file_table,
          lv_rc TYPE i,
          ls_record TYPE zst_all_job_ticket,
          lv_title TYPE string.

    FIELD-SYMBOLS:<l_fs> TYPE ANY.

    lv_title = text-004.

    " let the user to select the excel file  from application server
    CALL METHOD cl_gui_frontend_services=>file_open_dialog
      EXPORTING
        window_title            =  lv_title
      CHANGING
        file_table              = lt_filetable
        rc                      = lv_rc
*            user_action             =
*            file_encoding           =
      EXCEPTIONS
        file_open_dialog_failed = 1
        cntl_error              = 2
        error_no_gui            = 3
        not_supported_by_gui    = 4
        OTHERS                  = 5
            .
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                 WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    READ TABLE lt_filetable INDEX 1 INTO p_file .

    " convert the excel file data to internal table - raw
    " Since excel file has first line as header, we will consider from 2nd line
    CALL FUNCTION 'ALSM_EXCEL_TO_INTERNAL_TABLE'
      EXPORTING
        filename                = p_file
        i_begin_col             = 1
        i_begin_row             = 2
        i_end_col               = 29
        i_end_row               = 9999
      TABLES
        intern                  = lt_excel
      EXCEPTIONS
        inconsistent_parameters = 1
        upload_ole              = 2
        OTHERS                  = 3.
    IF sy-subrc = 0.
      SORT lt_excel BY row col.
      " now convert this data to required internal table (for uploading planned order) structure
      LOOP AT lt_excel INTO ls_excel.
        ASSIGN COMPONENT ls_excel-col OF STRUCTURE ls_record TO <l_fs>.
        MOVE ls_excel-value TO <l_fs>.
        AT END OF row.
          APPEND ls_record TO gt_excel_final.
          CLEAR ls_record.
        ENDAT.
      ENDLOOP.
      CLEAR :ls_excel,lt_excel.
    ENDIF.
  ENDMETHOD.                    "read_excel_file

  METHOD read_file_path.
    TYPES: BEGIN OF ty_header,
      werks TYPE char5,
      jttyp TYPE char25,
      sdate TYPE char25,
      empid TYPE char25,
      jshft TYPE char25,
      shdur TYPE char25,
      idtkt TYPE char25,
      sauft TYPE char40,
      matnr TYPE char25,
      verid TYPE char25,
      vornr TYPE char25,
      wcode TYPE char25,
      menge TYPE char25,
      xmnga TYPE char25,
      grund TYPE char25,
      stime TYPE char25,
      etime TYPE char15,
      durat TYPE char15,
      aufnr TYPE char15,
      arbpl TYPE char15,
      lstar TYPE char15,
      uebto TYPE char35,
      vlsch TYPE char35,
      zgroup TYPE char5,
      bereich TYPE char15,
      cavity TYPE char13,
      heats TYPE char13,
      delet TYPE char13,
     remarks TYPE char40,
           END OF ty_header.
    DATA: lv_title TYPE string,
          ls_path TYPE string.
    DATA: lv_fullpath TYPE string,
          lt_header TYPE TABLE OF ty_header,
          ls_header TYPE ty_header.
    lv_title = text-005.
    CLEAR ls_path.
    " let the user to select the location of the excel template to be downloaded
    " in the  application server
    CALL METHOD cl_gui_frontend_services=>directory_browse
      EXPORTING
        window_title    = lv_title"'File Directory'
        initial_folder  = 'C:'
      CHANGING
        selected_folder = ls_path.
    p_dfile = ls_path.

    ls_header-werks = text-006.
    ls_header-jttyp = text-007.
    ls_header-sdate = text-008.
    ls_header-empid = text-009.
    ls_header-jshft = text-010.
    ls_header-shdur = text-011.
    ls_header-idtkt = text-012.
    ls_header-sauft = text-013.
    ls_header-matnr = text-014.
    ls_header-verid = text-015.
    ls_header-vornr = text-016.
    ls_header-wcode = text-017.
    ls_header-menge = text-018.
    ls_header-xmnga = text-019.
    ls_header-grund = text-020.
    ls_header-stime = text-021.
    ls_header-etime = text-022.
    ls_header-durat = text-023.
    ls_header-aufnr = text-024.
    ls_header-arbpl = text-025.
    ls_header-lstar = text-026.
    ls_header-uebto = text-027.
    ls_header-vlsch = text-028.
    ls_header-zgroup = text-029.
    ls_header-bereich = text-030.
    ls_header-cavity = text-031.
    ls_header-heats = text-032.
    ls_header-delet = text-033.
    ls_header-remarks = text-034.
    APPEND ls_header TO lt_header.

    CONCATENATE ls_path '/' 'Excel Template'(035) sy-datum sy-uzeit '.XLS' INTO lv_fullpath.

    " Excel file template is downloaded in the selected location.
    CALL METHOD cl_gui_frontend_services=>gui_download
      EXPORTING
        filename              = lv_fullpath
        write_field_separator = 'X'
      CHANGING
        data_tab              = lt_header.

    CLEAR: lt_header, p_dfile.
  ENDMETHOD.                    "read_file_path


  METHOD get_instance.
    IF lo_selection IS NOT BOUND.
      CREATE OBJECT lo_selection.
    ENDIF.
    ro_instance = lo_selection.
  ENDMETHOD.                    "get_instance
ENDCLASS.                    "lcl_selscr IMPLEMENTATION

*----------------------------------------------------------------------*
*       CLASS lcl_ftp IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_ftp IMPLEMENTATION.

  METHOD ftp_cmd.
    CALL FUNCTION 'FTP_COMMAND'
      EXPORTING
        handle        = lv_ftp_handle
        command       = gv_cmd
        compress      = 'N'
      TABLES
        data          = lt_result
      EXCEPTIONS
        command_error = 1
        tcpip_error   = 2.
  ENDMETHOD.                    "ftp_cmd

  METHOD ftp_disconnect.
    CALL FUNCTION 'FTP_DISCONNECT'
      EXPORTING
        handle = lv_ftp_handle.
  ENDMETHOD.                    "ftp_disconnect

  METHOD rfc_connect.
    CALL FUNCTION 'RFC_CONNECTION_CLOSE'
      EXPORTING
        destination = i_rfc_destination
      EXCEPTIONS
        OTHERS      = 1.
  ENDMETHOD.                    "ftp_disconnect

  METHOD progress_indicator.
    CALL FUNCTION 'SAPGUI_PROGRESS_INDICATOR'
      EXPORTING
        text = text.
  ENDMETHOD.                    "process_indicator
  METHOD connect_ftp.
    DATA : i_password(30)     TYPE c,
           i_user(30)         TYPE c,
           i_host(30)         TYPE c,
           i_length           TYPE i,
           i_folder_path(100) TYPE c,
           i_filename(50)     TYPE c,
           i_efilename(50)     TYPE c,
           i_sfilename(50)     TYPE c.

    DATA:   lv_blob_length   TYPE i.
    DATA:   lv_length        TYPE i,  "Password length
            lv_key           TYPE i VALUE 26101957,
            lv_password(30)  TYPE c.

    TYPES: BEGIN OF ty_text,
           line(10000) TYPE c,
         END   OF ty_text.

    TYPES: BEGIN OF ty_status,
           file(50)     TYPE c,
           line         TYPE i,
           message TYPE bapi_msg,
           END OF ty_status.

    DATA: lt_text TYPE TABLE OF ty_text,
          ls_text LIKE LINE  OF lt_text,
          lt_status TYPE TABLE OF ty_status,
          ls_status TYPE ty_status.

    DATA: lt_blob_data  TYPE truxs_t_text_data,
          lt_error TYPE TABLE OF ty_text,
          lt_success TYPE TABLE OF ty_text,
          ls_blob TYPE ty_text,
          ls_header TYPE ty_text,
          g_sep(1) TYPE c VALUE '|',
          lv_str TYPE i,
          lv_shdur(8) TYPE c,
          lv_menge(16) TYPE c,
          lv_xmnga(16) TYPE c,
          lv_uebto(5)  TYPE c,
          lv_cavity(16) TYPE c,
          lv_heat(16) TYPE c,
          lv_dummy TYPE c,
          lv_eflag TYPE char1,
          lv_sflag TYPE char1,
          lv_remark(250) TYPE c,
          lv_s_no(30) TYPE c,
          lv_success TYPE c,
          lv_error TYPE c,
          lv_csv TYPE c,
          lv_finish TYPE c,
          lv_tabix  TYPE sy-tabix.

*----------------------*
    " Login Details of the FTP Server
    i_password        = 'Tristone_TF'.
    i_user            = 'tf'.
    i_host            = '10.58.146.64'.

    IF sy-batch = 'X'.
      i_rfc_destination = 'SAPFTPA'. "Runs in the Application Server
    ELSE.
      i_rfc_destination = 'SAPFTP'. " Runs in the Presentation Server
    ENDIF.

    i_length          = '992'.
    i_folder_path     = ''.

    " Header for Error and success table
    CONCATENATE text-006 text-007 text-008 text-009 text-010 text-011 text-012 text-013
    text-014 text-015 text-016 text-017 text-018 text-019 text-020 text-021 text-022 text-023
    text-024 text-025 text-026 text-027 text-028 text-029 text-030 text-031 text-032 text-033
    text-034 text-037 text-036 INTO ls_header SEPARATED BY '|'.

    APPEND ls_header TO lt_error.
    APPEND ls_header TO lt_success.
* Connect to server

    " Encrption of the FTP login details
    lv_length = STRLEN( i_password ).

    CALL FUNCTION 'HTTP_SCRAMBLE'
      EXPORTING
        SOURCE      = i_password
        sourcelen   = lv_length
        key         = lv_key
      IMPORTING
        destination = lv_password.

    "Establishment of FTP Server connection

    CALL FUNCTION 'FTP_CONNECT'
      EXPORTING
        user            = i_user
        password        = lv_password
        host            = i_host
        rfc_destination = i_rfc_destination
      IMPORTING
        handle          = lv_ftp_handle
      EXCEPTIONS
        not_connected   = 1
        OTHERS          = 2.
*Search file on FTP server
    me->progress_indicator( EXPORTING text = 'Select file on FTP Server'(039) ).
*  refresh result.
    " Activation of the Passive mode in the FTP Server, Passive mode acts as a firewall in the FTP instead of SSL.
    me->ftp_cmd( EXPORTING gv_cmd = 'set passive on'(040) ).

    " Input Folder is accessed

    CONCATENATE 'cd' '/Input' INTO lv_cmd SEPARATED BY space.

    me->ftp_cmd( EXPORTING gv_cmd = lv_cmd ).
    CLEAR lv_cmd.

    " LS command will retrive all the files available in the Input folder
    CONCATENATE 'ls' '' INTO lv_cmd SEPARATED BY space.
    me->ftp_cmd( EXPORTING gv_cmd = lv_cmd ).
    CLEAR lv_cmd.
    result = lt_result.
    LOOP AT result INTO ls_result.
      lv_tabix = sy-tabix.
      " Checks is the line contains .CSV File
      SEARCH ls_result-line FOR '.CSV'.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.
      " In FTP Passive mode Files are allowed to access only once,
      "Therefore for the second file the FTP is disconnected and reconnected again.
      IF lv_tabix > 1.
        CALL FUNCTION 'FTP_CONNECT'
          EXPORTING
            user            = i_user
            password        = lv_password
            host            = i_host
            rfc_destination = i_rfc_destination
          IMPORTING
            handle          = lv_ftp_handle
          EXCEPTIONS
            not_connected   = 1
            OTHERS          = 2.

        lv_cmd = 'set passive on'(040).
        me->ftp_cmd( EXPORTING gv_cmd = lv_cmd ).

        CONCATENATE 'cd' '/Input' INTO lv_cmd SEPARATED BY space.
        me->ftp_cmd( EXPORTING gv_cmd = lv_cmd ).
        CLEAR: lv_cmd, lt_result.
      ENDIF.
      " File Name format will be YYYYMMDDHHMM.CSV therefore last 16 characters are identified.
      lv_str = STRLEN( ls_result-line ).
      lv_str = lv_str - 16.
      i_filename = ls_result-line+lv_str(16).
      TRANSLATE i_filename TO LOWER CASE.
      "Information are captured from the specified file name.
      CALL FUNCTION 'FTP_SERVER_TO_R3'
        EXPORTING
          handle         = lv_ftp_handle
          fname          = i_filename
          character_mode = 'X'
        IMPORTING
          blob_length    = lv_blob_length
        TABLES
          blob           = lt_blob_data
          text           = lt_text.

      me->progress_indicator( EXPORTING text = 'Fetching the Contents from the file'(041) ).
*command for list the files from the directory
      "Captured data is retrived and saved in the internal table for Job ticket processing
      LOOP AT lt_text INTO ls_text FROM 2.
        SPLIT ls_text AT g_sep INTO gs_record-werks gs_record-jttyp
            gs_record-sdate gs_record-empid gs_record-jshft lv_shdur
            gs_record-idtkt gs_record-sauft gs_record-matnr gs_record-verid
            gs_record-vornr gs_record-wcode lv_menge lv_xmnga
            gs_record-grund gs_record-stime gs_record-etime gs_record-durat_d
            gs_record-aufnr gs_record-arbpl gs_record-lstar lv_uebto
            gs_record-vlsch gs_record-zgroup gs_record-bereich lv_cavity
            lv_heat gs_record-delet gs_record-remarks lv_dummy.

        gs_record-shdur = lv_shdur.
        gs_record-menge = lv_menge.
        gs_record-xmnga = lv_xmnga.
        gs_record-uebto = lv_uebto.
        gs_record-cavity = lv_cavity.
        gs_record-heats = lv_heat.
        APPEND gs_record TO go_selscr->gt_excel_final.
      ENDLOOP.

      IF go_selscr->gt_excel_final IS NOT INITIAL.
        lv_csv = 'X'.
      ENDIF.

      IF go_selscr->gt_excel_final IS NOT INITIAL.
        " Job Ticket Creation Process
        go_data = lcl_data=>get_instance( ).
        go_data->process_data( ).
      ELSE.

        CONCATENATE 'delete' i_filename INTO lv_cmd SEPARATED BY space.

        me->ftp_cmd( EXPORTING gv_cmd = lv_cmd ).
        CLEAR lv_cmd.

        " Error Folder is Opened
        CONCATENATE 'cd' '/Error' INTO lv_cmd SEPARATED BY space.

        me->ftp_cmd( EXPORTING gv_cmd = lv_cmd ).
        CLEAR lv_cmd.

        CONCATENATE 'No data Found' '' INTO ls_blob.
        APPEND ls_blob TO lt_error.

        ls_status-file = i_filename(12).
        ls_status-line = '0'.
        ls_status-message = 'No data Found'.
        APPEND ls_status TO lt_status.

        " CSV file is saved in the  FTP Folder
        CONCATENATE 'E' i_filename(12) '-' sy-uzeit '.CSV' INTO i_efilename.
*          TRANSLATE i_efilename TO LOWER CASE.
        CALL FUNCTION 'FTP_R3_TO_SERVER'
          EXPORTING
            handle         = lv_ftp_handle
            fname          = i_efilename
            character_mode = 'X'
          TABLES
            text           = lt_error
          EXCEPTIONS
            tcpip_error    = 1
            command_error  = 2
            data_error     = 3
            OTHERS         = 4.
        IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
          lv_error = 'X'.
        ENDIF.
        CLEAR: ls_status, ls_blob.
      ENDIF.

      IF go_selscr->gt_excel_final IS NOT INITIAL.
        " Captured feedback based on the Job ticket processing is separted based on the Success and Error Message
        IF go_data->gt_return IS NOT INITIAL.

          LOOP AT go_selscr->gt_excel_final INTO gs_record.
            READ TABLE go_data->gt_return INTO go_data->gs_return INDEX sy-tabix.
            IF sy-subrc = 0.
              lv_s_no = go_data->gs_return-s_no.
              CONDENSE lv_s_no.
              CONCATENATE   'Line'(050) lv_s_no INTO lv_remark
                            SEPARATED BY '-'.
              lv_shdur = gs_record-shdur.
              lv_menge = gs_record-menge.
              lv_xmnga = gs_record-xmnga.
              lv_uebto = gs_record-uebto.
              lv_cavity = gs_record-cavity.
              lv_heat = gs_record-heats.

              CONCATENATE gs_record-werks gs_record-jttyp
                gs_record-sdate gs_record-empid gs_record-jshft lv_shdur
                gs_record-idtkt gs_record-sauft gs_record-matnr gs_record-verid
                gs_record-vornr gs_record-wcode lv_menge lv_xmnga
                gs_record-grund gs_record-stime gs_record-etime gs_record-durat_d
                gs_record-aufnr gs_record-arbpl gs_record-lstar lv_uebto
                gs_record-vlsch gs_record-zgroup gs_record-bereich lv_cavity
                lv_heat gs_record-delet gs_record-remarks go_data->gs_return-message
                lv_remark INTO ls_blob SEPARATED BY '|'.

              ls_status-file = i_filename(12).
              ls_status-line = go_data->gs_return-s_no.
              ls_status-message = go_data->gs_return-message.
              APPEND ls_status TO lt_status.

              IF go_data->gs_return-type = 'E'.
                APPEND ls_blob TO lt_error.
                lv_eflag = 'X'.
              ELSE.
                APPEND ls_blob TO lt_success.
                lv_sflag = 'X'.
              ENDIF.

              CLEAR: lv_remark, ls_blob.
            ENDIF.
            CLEAR ls_status.
          ENDLOOP.

          " If Error Table contains values, then those details will be saved as a CSV file in error folder in FTP.
          IF lv_eflag = 'X'.
            " Error Folder is Opened
            CONCATENATE 'cd' '/Error' INTO lv_cmd SEPARATED BY space.
            me->ftp_cmd( EXPORTING gv_cmd = lv_cmd ).
            CLEAR lv_cmd.
            " CSV file is saved in the  FTP Folder
            CONCATENATE 'E' i_filename(12) '-' sy-uzeit '.CSV' INTO i_efilename.
*          TRANSLATE i_efilename TO LOWER CASE.
            CALL FUNCTION 'FTP_R3_TO_SERVER'
              EXPORTING
                handle         = lv_ftp_handle
                fname          = i_efilename
                character_mode = 'X'
              TABLES
                text           = lt_error
              EXCEPTIONS
                tcpip_error    = 1
                command_error  = 2
                data_error     = 3
                OTHERS         = 4.
            IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
              lv_error = 'X'.
            ENDIF.
            IF sy-subrc = 0.
              me->progress_indicator( EXPORTING text = 'Creation of File in Error Folder'(042) ).
            ENDIF.
          ENDIF.

          " If Success Table contains values, then those details will be saved as a CSV file in Success folder in FTP.

          IF lv_sflag = 'X'.
            " Success Folder is Opened
            CONCATENATE 'cd' '\Success' INTO lv_cmd SEPARATED BY space.

            me->ftp_cmd( EXPORTING gv_cmd = lv_cmd ).
            CLEAR lv_cmd.

            " CSV file is saved in the  FTP Folder
            CONCATENATE 'S' i_filename(12) '-' sy-uzeit '.CSV' INTO i_sfilename.
            CALL FUNCTION 'FTP_R3_TO_SERVER'
              EXPORTING
                handle         = lv_ftp_handle
                fname          = i_sfilename
                character_mode = 'X'
              TABLES
                text           = lt_success
              EXCEPTIONS
                tcpip_error    = 1
                command_error  = 2
                data_error     = 3
                OTHERS         = 4.
            IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
              lv_success = 'X'.
            ENDIF.

            IF sy-subrc = 0.
              me->progress_indicator( EXPORTING text = 'Creation of File in Success Folder'(043) ).
            ENDIF.
          ENDIF.

          " Input Folder is Opened
          CONCATENATE 'cd' '\Input' INTO lv_cmd SEPARATED BY space.

          me->ftp_cmd( EXPORTING gv_cmd = lv_cmd ).
          IF lv_success IS INITIAL AND lv_error IS INITIAL.
*        " Delete the existing file in the Input folder
            CONCATENATE 'delete' i_filename INTO lv_cmd SEPARATED BY space.

            go_ftp->ftp_cmd( EXPORTING gv_cmd = lv_cmd ).
            IF sy-subrc = 0.
              lv_finish = 'X'.
              me->progress_indicator( EXPORTING text = 'Deletion of File in Input Folder'(044) ).
            ENDIF.
          ENDIF.
        ENDIF.
      ENDIF.
      CLEAR: lt_text, lt_result,lv_cmd,lv_ftp_handle,i_sfilename, i_efilename, i_filename, go_selscr->gt_excel_final.

      me->ftp_disconnect( ).

      " RFC Coonection is closed
      me->rfc_connect( ).
    ENDLOOP.
    CLEAR result.

    " FTP is disconnected
    me->ftp_disconnect( ).
    me->progress_indicator( EXPORTING text = 'FTP Getting Disconnecetd'(045) ).
    " RFC Connection is closed
    me->rfc_connect( ).
    me->progress_indicator( EXPORTING text = 'RFC Getting Disconnecetd'(046) ).
    IF lv_csv IS INITIAL.
      WRITE: 'No Input File is available in the Input Folder of the FTP Server for processing'(048).
    ENDIF.

    IF lv_finish  = 'X'.
      WRITE: 'Job Ticket Processed'(047) COLOR 2.
      WRITE:/ 'FileName  |                   Lineno |' COLOR 3.
      WRITE:/ 'Status' COLOR 3.
      LOOP AT lt_status INTO ls_status.
        CONDENSE ls_status-message.
        CONDENSE ls_status-file.
        WRITE:/ ls_status-file, ls_status-line, ls_status-message NO-GAP.
      ENDLOOP.
    ENDIF.
    CLEAR lt_status.
  ENDMETHOD.                    "connect_ftp

  METHOD get_instance.
    IF lo_ftp IS NOT BOUND.
      CREATE OBJECT lo_ftp.
    ENDIF.
    ro_instance = lo_ftp.
  ENDMETHOD.                    "get_instance
ENDCLASS.                    "lcl_ftp IMPLEMENTATION

*----------------------------------------------------------------------*
*       CLASS lcl_data IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_data IMPLEMENTATION.
  METHOD process_data.
    LOOP AT go_selscr->gt_excel_final INTO gs_final.
      lv_count = lv_count + 1.
      " Job ticket creation is performed using the data given in th excel file
      CALL FUNCTION 'ZPP_CREATE_JOB_TICKET'
        EXPORTING
          is_job      = gs_final
        TABLES
          return      = lt_return
        EXCEPTIONS
          error_found = 1
          OTHERS      = 2.
      IF sy-subrc <> 0.
        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4 INTO ls_return1.
      ENDIF.
      LOOP AT lt_return INTO ls_return.
        MOVE-CORRESPONDING ls_return TO gs_return.
        gs_return-s_no = lv_count.
        APPEND gs_return TO gt_return.
      ENDLOOP.
      CLEAR lt_return.
    ENDLOOP.
  ENDMETHOD.                    "process_data

  METHOD get_instance.
    IF lo_data IS NOT BOUND.
      CREATE OBJECT lo_data.
    ENDIF.
    ro_instance = lo_data.
  ENDMETHOD.                    "get_instance
ENDCLASS.                    "lcl_data IMPLEMENTATION

*----------------------------------------------------------------------*
*       CLASS lcl_alv IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_alv IMPLEMENTATION.
  METHOD display_alv.
    me->salv_factory( IMPORTING eo_salv = lo_salv CHANGING ctab = go_data->gt_return ).
    me->set_functions( CHANGING co_salv = lo_salv ).
    me->set_zebra( CHANGING co_salv = lo_salv ).
    me->set_column_optimization( CHANGING co_salv = lo_salv ).
    me->set_column_text( EXPORTING iv_ltext = 'Excel_Row_No'  iv_column_name = 'S_NO'
                                   iv_mtext = 'Excel_Row_No' "iv_stext = 'Emp. Name'
                     	    CHANGING co_salv = lo_salv ).
    lo_salv->display( ).
  ENDMETHOD.                    "Display_alv

  METHOD get_instance.
    IF lo_alv IS NOT BOUND.
      CREATE OBJECT lo_alv.
    ENDIF.
    ro_instance = lo_alv.
  ENDMETHOD.                    "get_instance
ENDCLASS.                    "lcl_alv IMPLEMENTATION
*-------------------------------- End of Class Implementations --------------------------------&

*----------------------------------- Begin of Event Blocks ------------------------------------&

INITIALIZATION.
  go_selscr = lcl_selscr=>get_instance( ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  go_selscr->read_excel_file( ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_dfile.
  go_selscr->read_file_path( ).

AT SELECTION-SCREEN OUTPUT.

  IF p_upload = 'X'.
    LOOP AT SCREEN.
      IF screen-group1 = 'FL'.
        screen-active = 1.
      ENDIF.
      IF screen-group1 = 'DL'.
        screen-active = 0.
      ENDIF.
      MODIFY SCREEN.
    ENDLOOP.
  ELSEIF p_downld = 'X'.
    LOOP AT SCREEN.
      IF screen-group1 = 'FL'.
        screen-active = 0.
      ENDIF.
      IF screen-group1 = 'DL'.
        screen-active = 1.
      ENDIF.
      MODIFY SCREEN.
    ENDLOOP.
  ELSEIF p_backg = 'X'.
    LOOP AT SCREEN.
      IF screen-group1 = 'FL'.
        screen-active = 0.
      ENDIF.
      IF screen-group1 = 'DL'.
        screen-active = 0.
      ENDIF.
      MODIFY SCREEN.
    ENDLOOP.
  ENDIF.

AT SELECTION-SCREEN .
  IF go_selscr->gt_excel_final IS INITIAL AND p_upload = 'X' AND sy-ucomm = 'ONLI' .
    MESSAGE e004. "'Excel Upload Failed' TYPE 'E'.
  ENDIF.

START-OF-SELECTION.
  CLEAR p_file.

  IF p_backg = 'X'.
    go_ftp = lcl_ftp=>get_instance( ).
    go_ftp->connect_ftp( ).
  ENDIF.
  IF p_backg <> 'X'.
    go_data = lcl_data=>get_instance( ).
    go_data->process_data( ).
    go_alv = lcl_alv=>get_instance( ).
    go_alv->display_alv( ).
  ENDIF.
