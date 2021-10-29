*&---------------------------------------------------------------------*
*& Report  ZABAP_TBTCO_SM37_LOG
*&
*&---------------------------------------------------------------------*
*& SM37 but showing log messages, has also BW filters and WPER logs
*&---------------------------------------------------------------------*
report  zabap_tbtco_sm37_log.

tables: tbtco, tbtcp, rspcvariant, rsparams.

types: begin of t_styles,
         header_up      type zexcel_cell_style,
         header_edit    type zexcel_cell_style,
         header_down    type zexcel_cell_style,
         normal         type zexcel_cell_style,
         edit           type zexcel_cell_style,
         red            type zexcel_cell_style,
         light_red      type zexcel_cell_style,
         green          type zexcel_cell_style,
         yellow         type zexcel_cell_style,
       end of t_styles.

types: begin of t_job,
         jobname   type tbtco-jobname,
         sdlstrtdt type tbtco-sdlstrtdt,
         sdlstrttm type tbtco-sdlstrttm,
         sdluname  type tbtco-sdluname,
         status    type zbtc_status,

         entertime type tbtc5-entertime,
         enterdate type tbtc5-enterdate,
         msgid     type tbtc5-msgid,
         msgno     type tbtc5-msgno,
         msgtype   type tbtc5-msgtype,
         text      type tbtc5-text,

         jobcount type tbtco-jobcount,
         errkey(23),

         stepcount type tbtcp-stepcount,
         progname type tbtcp-progname,
         variant type tbtcp-variant,
         authcknam type tbtcp-authcknam,
         bw_chain type rsparams-low,
         bw_step type rspcvariant-low,
         bw_program type rspcvariant-low,
         bw_variant type rspcvariant-low,
         bw_destination type rspcvariant-low,

         colortab type lvc_t_scol,

       end of t_job.
data: gt_jobs type table of t_job.
data: gs_job type t_job.
data: gr_salv type ref to cl_salv_table.
data: gr_repid type range of rspcvariant-low.

select-options: s_jobnam for tbtco-jobname,
                s_strtdt for tbtco-sdlstrtdt,
                s_strttm for tbtco-sdlstrttm,
                s_uname for tbtco-sdluname matchcode object user_comp,
                s_status for gs_job-status,
                s_errkey for gs_job-errkey,
                s_msgty for syst-msgty.
selection-screen begin of block c2 with frame title text-005.
parameters: p_steps as checkbox.
select-options: s_cprog for sy-cprog.
select-options: s_chain for rsparams-low.
select-options: s_step for rspcvariant-low.
select-options: s_repid for sy-repid.
select-options: s_vari for rspcvariant-low.
selection-screen end of block c2.

selection-screen begin of block b4 with frame title text-015.
parameters p_var type slis_vari.
selection-screen end of block b4.

*  btc_unknown_state like tbtco-status value 'X'.
selection-screen begin of block c1 with frame title text-004.
select-options: s_errkyr for gs_job-errkey.
selection-screen end of block c1.
selection-screen begin of block b1 with frame title text-001.
parameters: p_sender type adr6-smtp_addr default 'no@spam.org' obligatory.
parameters: p_title type so_obj_des.
selection-screen begin of block b2 with frame title text-002.
parameters: p_email  type adr6-smtp_addr.
selection-screen end of block b2.
selection-screen begin of block b3 with frame title text-003.
parameters: p_attach as checkbox.
selection-screen end of block b3.
selection-screen end of block b1.

*----------------------------------------------------------------------*
*       CLASS lcl_event_handle DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class lcl_event_handle definition.
  public section.
    methods : hndl_double_click for event double_click of cl_salv_events_table importing row column.
endclass.                    "lcl_event_handle DEFINITION
*----------------------------------------------------------------------*
*       CLASS lcl_event_handle IMPLEMENTATION
*----------------------------------------------------------------------*
class lcl_event_handle implementation.
  method hndl_double_click.
    field-symbols <ls_job> type t_job.

    read table gt_jobs assigning <ls_job> index row.
    if sy-subrc eq 0.
      perform f_show_job using <ls_job>.
    endif.
  endmethod.                    "hndl_double_click
endclass.                    "lcl_event_handle IMPLEMENTATION

initialization.
  s_strtdt-sign = 'I'.
  s_strtdt-option = 'EQ'.
  s_strtdt-low = sy-datum - 1.
  append s_strtdt.

  s_msgty-sign = 'I'.
  s_msgty-option = 'EQ'.
  s_msgty-low = 'E'.
  append s_msgty.
  s_msgty-sign = 'I'.
  s_msgty-option = 'EQ'.
  s_msgty-low = 'A'.
  append s_msgty.
  s_msgty-sign = 'I'.
  s_msgty-option = 'EQ'.
  s_msgty-low = 'X'.
  append s_msgty.

at selection-screen on value-request for p_var.
  data: ls_variant type disvariant,
        lv_exit    type char1.

  ls_variant-report = sy-repid.

  call function 'LVC_VARIANT_F4'
    exporting
      is_variant = ls_variant
      i_save     = 'A'
    importing
      e_exit     = lv_exit
      es_variant = ls_variant
    exceptions
      not_found  = 1
      others     = 2.
  if sy-subrc ne 0.
    message id sy-msgid type sy-msgty number sy-msgno
            with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  else.
    check lv_exit eq space.
    p_var = ls_variant-variant.
  endif.

start-of-selection.
  perform f_set_r_repid.
  perform f_select.


end-of-selection.
  perform f_set_colors.
  if p_email is not initial.
    perform f_send_email.
  else.
    perform f_display.
  endif.


*&---------------------------------------------------------------------*
*&      Form  F_SHOW_JOB
*&---------------------------------------------------------------------*
form f_show_job  using    ps_job type t_job.
  data: lr_btci type ref to zbatch_input,
        l_date(10),
        l_time(8).

  create object lr_btci.

  lr_btci->bdc_dynpro( program = 'SAPLBTCH' dynpro  = '2170' ).
  lr_btci->bdc_field( fnam   = 'BDC_OKCODE' fval   = '=DOIT' ).
  lr_btci->bdc_field( fnam   = 'BTCH2170-JOBNAME' fval   = ps_job-jobname ).
  lr_btci->bdc_field( fnam   = 'BTCH2170-USERNAME' fval   = ps_job-sdluname ).
  lr_btci->bdc_field( fnam   = 'BTCH2170-SCHEDUL' fval   = '' ).
  lr_btci->bdc_field( fnam   = 'BTCH2170-READY' fval   = '' ).
  lr_btci->bdc_field( fnam   = 'BTCH2170-RUNNING' fval   = 'X' ).
  lr_btci->bdc_field( fnam   = 'BTCH2170-FINISHED' fval   = 'X' ).
  lr_btci->bdc_field( fnam   = 'BTCH2170-ABORTED' fval   = 'X' ).
  write ps_job-sdlstrtdt to l_date.
  write ps_job-sdlstrttm to l_time.
*  write ps_job-enterdate to l_date.
*  write ps_job-entertime to l_time.
  lr_btci->bdc_field( fnam   = 'BTCH2170-FROM_DATE' fval   = l_date ).
  lr_btci->bdc_field( fnam   = 'BTCH2170-FROM_TIME' fval   = l_time ).
  lr_btci->bdc_field( fnam   = 'BTCH2170-TO_DATE' fval   = l_date ).
  lr_btci->bdc_field( fnam   = 'BTCH2170-TO_TIME' fval   = l_time ).

*  lr_btci->bdc_dynpro( program = 'SAPMSSY0' dynpro  = '0120' ).
*  lr_btci->bdc_field( fnam   = 'BDC_OKCODE' fval   = '=BKWD' ).
*
  lr_btci->bdc_dynpro( program = 'SAPLBTCH' dynpro  = '2170' ).
  lr_btci->bdc_field( fnam   = 'BDC_OKCODE' fval   = '/EECAN' ).

  lr_btci->execute( tcode   = 'SM37' ).

endform.                    " F_SHOW_JOB
*&---------------------------------------------------------------------*
*&      Form  F_SELECT
*&---------------------------------------------------------------------*
form f_select .
  data: lt_log type table of tbtc5.
  field-symbols: <ls_log> type tbtc5.
  field-symbols: <ls_job> type t_job.
  data: lt_wrpe type table of wrpe.
  field-symbols <ls_wrpe> type wrpe.
  data: l_rplid type wrpe-rplid.
  data: lt_tbtcp type table of tbtcp.
  data: ls_tbtcp type tbtcp.
  data: begin of ls_bw,
          bw_chain type rsparams-low,
          bw_step type rspcvariant-low,
          bw_program type rspcvariant-low,
          bw_variant type rspcvariant-low,
          bw_destination type rspcvariant-low,
        end of ls_bw.
  data: lt_value type table of rsparams.
  field-symbols: <ls_value> type rsparams.


  select * from tbtco where jobname in s_jobnam and sdlstrtdt in s_strtdt and sdluname in s_uname and status in s_status and sdlstrttm in s_strttm.
    if p_steps is not initial.
      select * into table lt_tbtcp from tbtcp where jobname eq tbtco-jobname and jobcount eq tbtco-jobcount order by stepcount.
    endif.
    clear lt_log.
    clear: ls_bw, ls_tbtcp.
    call function 'BP_JOBLOG_READ'
      exporting
        client                = sy-mandt
        jobcount              = tbtco-jobcount
        joblog                = tbtco-joblog
        jobname               = tbtco-jobname
      tables
        joblogtbl             = lt_log
      exceptions
        cant_read_joblog      = 1
        jobcount_missing      = 2
        joblog_does_not_exist = 3
        joblog_is_empty       = 4
        joblog_name_missing   = 5
        jobname_missing       = 6
        job_does_not_exist    = 7
        others                = 8.
    if sy-subrc eq 0.
      loop at lt_log assigning <ls_log>. "WHERE msgtype EQ 'E' OR msgtype EQ 'A' OR msgtype EQ 'X'.


        if p_steps is not initial.
          if <ls_log>-msgtype eq 'S' and <ls_log>-msgid eq '00' and <ls_log>-msgno eq '550'.
            clear: ls_bw, ls_tbtcp.
            read table lt_tbtcp into ls_tbtcp index 1.
            if sy-subrc eq 0.
              delete lt_tbtcp index 1.
            endif.
            if ls_tbtcp-variant is not initial.
              clear lt_value.
              call function 'RS_VARIANT_CONTENTS'
                exporting
                  report               = ls_tbtcp-progname
                  variant              = ls_tbtcp-variant
                tables
                  valutab              = lt_value
                exceptions
                  variant_non_existent = 1
                  variant_obsolete     = 2
                  others               = 3.

              read table lt_value assigning <ls_value> with key selname = 'CHAIN'.
              if sy-subrc eq 0.
                ls_bw-bw_chain = <ls_value>-low.
              endif.
              read table lt_value assigning <ls_value> with key selname = 'VARIANT'.
              if sy-subrc eq 0.
                ls_bw-bw_step = <ls_value>-low.

                if ls_bw-bw_step is not initial.
                  select single low into ls_bw-bw_program from rspcvariant where type eq 'ABAP' and variante eq ls_bw-bw_step and objvers eq 'A' and fnam eq 'PROGRAM'.
                  select single low into ls_bw-bw_variant from rspcvariant where type eq 'ABAP' and variante eq ls_bw-bw_step and objvers eq 'A' and fnam eq 'VARIANT'.
                  select single low into ls_bw-bw_destination from rspcvariant where type eq 'REMOTE' and variante eq ls_bw-bw_step and objvers eq 'A' and fnam eq 'DEST_TARGET'.
                endif.
              endif.
            endif.
          elseif <ls_log>-msgtype eq 'S' and <ls_log>-msgid eq '00' and <ls_log>-msgno eq '517'.
            clear: ls_bw, ls_tbtcp.
          endif.
        endif.

        append initial line to gt_jobs assigning <ls_job>.
        move-corresponding tbtco to <ls_job>.
        if ls_tbtcp is not initial.
          move-corresponding ls_tbtcp to <ls_job>.
        endif.
        move-corresponding ls_bw to <ls_job>.
        move-corresponding <ls_log> to <ls_job>.
        concatenate <ls_job>-msgid <ls_job>-msgno into <ls_job>-errkey.

        if <ls_job>-msgid eq 'WT' and <ls_job>-msgno eq '213'. "log execução reabastecimento
          l_rplid = <ls_job>-text+28(10).
          select * into table lt_wrpe from wrpe where rplid eq l_rplid order by kunnr matnr seqnr.
          read table lt_wrpe with key msgty = 'E' transporting no fields.
          if sy-subrc eq 0.
            loop at lt_wrpe assigning <ls_wrpe>.
              at new matnr.
                append initial line to gt_jobs assigning <ls_job>.
                move-corresponding tbtco to <ls_job>.
                if ls_tbtcp is not initial.
                  move-corresponding ls_tbtcp to <ls_job>.
                endif.
                move-corresponding ls_bw to <ls_job>.
                move-corresponding <ls_log> to <ls_job>.
                <ls_job>-msgtype = 'I'.
                <ls_job>-msgid = 'ZFAIRSHARE'.
                <ls_job>-msgno = '002'.
                concatenate <ls_job>-msgid <ls_job>-msgno into <ls_job>-errkey.
                message id <ls_job>-msgid type <ls_job>-msgtype number <ls_job>-msgno with <ls_wrpe>-matnr <ls_wrpe>-kunnr into <ls_job>-text.
              endat.
              append initial line to gt_jobs assigning <ls_job>.
              move-corresponding tbtco to <ls_job>.
              if ls_tbtcp is not initial.
                move-corresponding ls_tbtcp to <ls_job>.
              endif.
              move-corresponding ls_bw to <ls_job>.
              move-corresponding <ls_log> to <ls_job>.
              <ls_job>-msgtype = <ls_wrpe>-msgty.
              <ls_job>-msgid = <ls_wrpe>-msgid.
              <ls_job>-msgno = <ls_wrpe>-msgno.
              concatenate <ls_job>-msgid <ls_job>-msgno into <ls_job>-errkey.
              message id <ls_job>-msgid type <ls_job>-msgtype number <ls_job>-msgno with <ls_wrpe>-msgv1 <ls_wrpe>-msgv2 <ls_wrpe>-msgv3 <ls_wrpe>-msgv4 into <ls_job>-text.
            endloop.
          endif.
        elseif <ls_job>-msgid eq 'MG' and <ls_job>-msgno eq '537'. "log materiais SLG1, nº log= msgv2, objecto = 'MATU'
          data: l_extnumber type balhdr-extnumber,
                lt_balm type table of balm.
          field-symbols <ls_balm> type balm.
          data: ls_result type match_result .

          clear lt_balm.
          find first occurrence of ': nº log' in <ls_job>-text results ls_result.
          if sy-subrc eq 0.
            add 9 to ls_result-offset.
            l_extnumber = <ls_job>-text+ls_result-offset(16).
            call function 'APPL_LOG_READ_DB'
              exporting
                object          = 'MATU'
                external_number = l_extnumber
              tables
                messages        = lt_balm.
          endif.
          loop at lt_balm assigning <ls_balm> where msgty in s_msgty.

            append initial line to gt_jobs assigning <ls_job>.
            move-corresponding tbtco to <ls_job>.
            if ls_tbtcp is not initial.
              move-corresponding ls_tbtcp to <ls_job>.
            endif.
            move-corresponding ls_bw to <ls_job>.
            move-corresponding <ls_log> to <ls_job>.
            move-corresponding <ls_balm> to <ls_job>.
            <ls_job>-msgtype = <ls_balm>-msgty.
            <ls_job>-msgid = <ls_balm>-msgid.
            <ls_job>-msgno = <ls_balm>-msgno.
            concatenate <ls_job>-msgid <ls_job>-msgno into <ls_job>-errkey.
            message id <ls_job>-msgid type <ls_job>-msgtype number <ls_job>-msgno with <ls_balm>-msgv1 <ls_balm>-msgv2 <ls_balm>-msgv3 <ls_balm>-msgv4 into <ls_job>-text.
          endloop.
        endif.

      endloop.
    endif.
    delete gt_jobs where msgtype not in s_msgty and not ( msgid eq 'WT' and msgno eq '213' ).
    delete gt_jobs where errkey not in s_errkey or msgtype not in s_msgty.
  endselect.
  if p_steps is not initial.
    delete gt_jobs where bw_chain not in s_chain
                     or bw_step not in s_step
                     or bw_program not in gr_repid
                     or bw_variant not in s_vari
                     or progname not in s_cprog.
  endif.

endform.                    " F_SELECT
*&---------------------------------------------------------------------*
*&      Form  F_DISPLAY
*&---------------------------------------------------------------------*
form f_display .
  cl_salv_table=>factory(
    importing
      r_salv_table = gr_salv
    changing
      t_table      = gt_jobs ).

  data: lr_functions type ref to cl_salv_functions_list.
  lr_functions = gr_salv->get_functions( ).
  lr_functions->set_all( abap_true ).
  lr_functions->set_default( abap_true ).
  data: lo_funcs type salv_t_ui_func.
  data: lo_func type salv_s_ui_func.
  lo_funcs = lr_functions->get_functions( ).
  loop at lo_funcs into lo_func.
    lo_func-r_function->set_enable( value = 'X' ).
    lo_func-r_function->set_visible( value = 'X' ).
  endloop.

  data: lo_event     type ref to cl_salv_events_table,
        lo_handle    type ref to lcl_event_handle.
  create object lo_handle.
  call method gr_salv->get_event
    receiving
      value = lo_event.    "Get all the Events of the table
  set handler lo_handle->hndl_double_click for lo_event.

  data: lo_layout  type ref to cl_salv_layout,
        lf_variant type slis_vari,
        ls_key    type salv_s_layout_key.
  lo_layout = gr_salv->get_layout( ).
  ls_key-report = sy-repid.
  lo_layout->set_key( ls_key ).
  lo_layout->set_save_restriction( if_salv_c_layout=>restrict_none ).
  lo_layout->set_default( abap_true ).

  lf_variant = p_var.

  lo_layout->set_initial_layout( lf_variant ).

  data: lo_column      type ref to cl_salv_column.
  data: lo_columns     type ref to cl_salv_columns_table.

  lo_columns = gr_salv->get_columns( ).

  lo_columns->set_color_column( 'COLORTAB' ).

  lo_column = lo_columns->get_column( 'ERRKEY' ).
  lo_column->set_short_text( 'Erro'(c21) ).
  lo_column->set_medium_text( 'Erro'(c21) ).
  lo_column->set_long_text( 'Erro'(c21) ).

  lo_column = lo_columns->get_column( 'BW_CHAIN' ).
  lo_column->set_short_text( 'Cadeia'(c22) ).
  lo_column->set_medium_text( 'Cadeia'(c22) ).
  lo_column->set_long_text( 'Cadeia'(c22) ).
  lo_column = lo_columns->get_column( 'BW_STEP' ).
  lo_column->set_short_text( 'Etapa'(c23) ).
  lo_column->set_medium_text( 'Etapa'(c23) ).
  lo_column->set_long_text( 'Etapa'(c23) ).
  lo_column = lo_columns->get_column( 'BW_PROGRAM' ).
  lo_column->set_short_text( 'ProgramaCa'(c24) ).
  lo_column->set_medium_text( 'Programa Cadeia'(c25) ).
  lo_column->set_long_text( 'Programa Cadeia'(c25) ).
  lo_column = lo_columns->get_column( 'BW_VARIANT' ).
  lo_column->set_short_text( 'VarianteCa'(c27) ).
  lo_column->set_medium_text( 'Variante Cadeia'(c28) ).
  lo_column->set_long_text( 'Variante Cadeia'(c27) ).
  lo_column = lo_columns->get_column( 'BW_DESTINATION' ).
  lo_column->set_short_text( 'DestinoCa'(c29) ).
  lo_column->set_medium_text( 'Destino Cadeia'(c26) ).
  lo_column->set_long_text( 'Destino Cadeia'(c26) ).

  gr_salv->display( ).

endform.                    " F_DISPLAY
*&---------------------------------------------------------------------*
*&      Form  F_SEND_EMAIL
*&---------------------------------------------------------------------*
form f_send_email .
  field-symbols <ls_job> type t_job.
  data: lo_excel type ref to zcl_excel.

  data: cl_writer type ref to zif_excel_writer.
  data: xdata       type xstring,             " Will be used for sending as email
        t_rawdata   type solix_tab,           " Will be used for downloading or open directly
        bytecount   type i.                   " Will be used for downloading or open directly
  data: l_save_file_name type string.
* Needed to send emails
  data: bcs_exception           type ref to cx_bcs,
        errortext               type string,
        cl_send_request         type ref to cl_bcs,
        cl_document             type ref to cl_document_bcs,
        cl_recipient            type ref to if_recipient_bcs,
        cl_sender               type ref to cl_cam_address_bcs,
        t_attachment_header     type soli_tab,
        wa_attachment_header    like line of t_attachment_header,
        attachment_subject      type sood-objdes,
        sood_bytecount          type sood-objlen,
        mail_title              type so_obj_des,
        t_mailtext              type soli_tab,
        wa_mailtext             like line of t_mailtext,
        send_to                 type adr6-smtp_addr,
        sent                    type os_boolean.
  data: lt_html          type bcsy_text,
        ls_html          type line of bcsy_text.
  data: l_field(255)." TYPE string.
  data: l_td type string.

  data: l_date(10).

  mail_title = p_title.
  read table s_strtdt index 1.
  write s_strtdt-low to l_date.
  replace 'S_STRTDT' in mail_title with l_date.

  read table s_jobnam index 1.
  replace 'S_JOBNAM' in mail_title with s_jobnam-low.

  replace 'SYSID' in mail_title with sy-sysid.

  concatenate mail_title '.xlsx' into l_save_file_name.

  append '<head>' to lt_html.
  append '<style type="text/css">' to lt_html.
  append '<!--' to lt_html.
  append 'p{font-family: verdana; font-size:x-small}' to lt_html.
  append 'table{background-color:#FFF;border-collapse:collapse; font-family: verdana; font-size:xx-small}' to lt_html.
  append 'th.Heading{background-color:#0F243E;text-align:left;color:white;border:1px solidblack;padding:3px;font-family: verdana; font-size:xx-small}' to lt_html. "text-align:center;
  append 'td.red{background-color:#FF3333;border:1px solid black;padding:3px;font-family: verdana; font-size:xx-small}' to lt_html.
  append 'td.light_red{background-color:#FFCCCC;border:1px solid black;padding:3px;font-family: verdana; font-size:xx-small}' to lt_html.
  append 'td.green{background-color:#00CC44;border:1px solid black;padding:3px;font-family: verdana; font-size:xx-small}' to lt_html.
  append 'td.yellow{background-color:#FFFF00;border:1px solid black;padding:3px;font-family: verdana; font-size:xx-small}' to lt_html.
  append 'td.normal{background-color:#FFF;border:1px solid black;padding:3px;font-family: verdana; font-size:xx-small}' to lt_html.

*  lo_style_red->fill->fgcolor-rgb  = 'FF3333'. "red
*  lo_style_light_red->fill->fgcolor-rgb  = 'FFCCCC'. "light_red
*  lo_style_green->fill->fgcolor-rgb  = '00CC44'. "green
*  lo_style_yellow->fill->fgcolor-rgb  = 'FFFF00'. "yellow
*  lo_style_edit->fill->fgcolor-rgb  = 'FFFF66'. "top zahara blue

  append '-->' to lt_html.
  append '</style>' to lt_html.
  append '</head>' to lt_html.
  append '<body>' to lt_html.
  append '<p>Bom dia,</p>' to lt_html.
  if gt_jobs is initial.
    append '<p>Nenhum log de jobs de acordo com critérios de seleção</p>'(h01) to lt_html.
  else.
    append '<p>Segue a listagem de erros de execução de jobs.</p>'(h02) to lt_html.
  endif.
  if p_attach is not initial.

    if gt_jobs is not initial.
      perform f_create_list_in_excel using lo_excel.

      create object cl_writer type zcl_excel_writer_2007.
      xdata = cl_writer->write_file( lo_excel ).
      t_rawdata = cl_bcs_convert=>xstring_to_solix( iv_xstring  = xdata ).
      bytecount = xstrlen( xdata ).
    endif.

  else.
    if gt_jobs is not initial.
      append '<p><table>' to lt_html.
      append '<tr>' to lt_html.
      append '<th class="Heading">Job</th>' to lt_html.
      append '<th class="Heading">Execução</th>' to lt_html.
      append '<th class="Heading">Execução</th>' to lt_html.
      append '<th class="Heading">Escalonador</th>' to lt_html.
      append '<th class="Heading">Status</th>' to lt_html.
      append '<th class="Heading">Log</th>' to lt_html.
      append '<th class="Heading">Log</th>' to lt_html.
      append '<th class="Heading">ID</th>' to lt_html.
      append '<th class="Heading">Nº</th>' to lt_html.
      append '<th class="Heading">Tipo</th>' to lt_html.
      append '<th class="Heading">Erro</th>' to lt_html.

      if p_steps is not initial.
        append '<th class="Heading">Cadeia</th>' to lt_html.
        append '<th class="Heading">Etapa</th>' to lt_html.
        append '<th class="Heading">Programa Cadeia</th>' to lt_html.
        append '<th class="Heading">Variante Cadeia</th>' to lt_html.
        append '<th class="Heading">Destino Cadeia</th>' to lt_html.
      endif.

      append '</tr>' to lt_html.

      loop at gt_jobs assigning <ls_job>.
        append '<tr>' to lt_html.
        perform f_get_td_class using <ls_job> 'JOBNAME' l_td.
        append l_td to lt_html. write <ls_job>-jobname to l_field. append l_field to lt_html. append '</td>' to lt_html.
        perform f_get_td_class using <ls_job> 'SDLSTRTDT' l_td.
        append l_td to lt_html. write <ls_job>-sdlstrtdt to l_field. append l_field to lt_html. append '</td>' to lt_html.
        perform f_get_td_class using <ls_job> 'SDLSTRTTM' l_td.
        append l_td to lt_html. write <ls_job>-sdlstrttm to l_field. append l_field to lt_html. append '</td>' to lt_html.
        perform f_get_td_class using <ls_job> 'SDLUNAME' l_td.
        append l_td to lt_html. write <ls_job>-sdluname to l_field. append l_field to lt_html. append '</td>' to lt_html.
        perform f_get_td_class using <ls_job> 'STATUS' l_td.
        append l_td to lt_html. write <ls_job>-status to l_field. append l_field to lt_html. append '</td>' to lt_html.
        perform f_get_td_class using <ls_job> 'ENTERDATE' l_td.
        append l_td to lt_html. write <ls_job>-enterdate to l_field. append l_field to lt_html. append '</td>' to lt_html.
        perform f_get_td_class using <ls_job> 'ENTERTIME' l_td.
        append l_td to lt_html. write <ls_job>-entertime to l_field. append l_field to lt_html. append '</td>' to lt_html.
        perform f_get_td_class using <ls_job> 'MSGID' l_td.
        append l_td to lt_html. write <ls_job>-msgid to l_field. append l_field to lt_html. append '</td>' to lt_html.
        perform f_get_td_class using <ls_job> 'MSGNO' l_td.
        append l_td to lt_html. write <ls_job>-msgno to l_field. append l_field to lt_html. append '</td>' to lt_html.
        perform f_get_td_class using <ls_job> 'MSGTYPE' l_td.
        append l_td to lt_html. write <ls_job>-msgtype to l_field. append l_field to lt_html. append '</td>' to lt_html.
        perform f_get_td_class using <ls_job> 'TEXT' l_td.
        append l_td to lt_html. write <ls_job>-text to l_field. append l_field to lt_html. append '</td>' to lt_html.

        if p_steps is not initial.
          perform f_get_td_class using <ls_job> 'BW_CHAIN' l_td.
          append l_td to lt_html. write <ls_job>-bw_chain to l_field. append l_field to lt_html. append '</td>' to lt_html.

          perform f_get_td_class using <ls_job> 'BW_STEP' l_td.
          append l_td to lt_html. write <ls_job>-bw_step to l_field. append l_field to lt_html. append '</td>' to lt_html.
          perform f_get_td_class using <ls_job> 'BW_PROGRAM' l_td.
          append l_td to lt_html. write <ls_job>-bw_program to l_field. append l_field to lt_html. append '</td>' to lt_html.
          perform f_get_td_class using <ls_job> 'BW_VARIANT' l_td.
          append l_td to lt_html. write <ls_job>-bw_variant to l_field. append l_field to lt_html. append '</td>' to lt_html.
          perform f_get_td_class using <ls_job> 'BW_DESTINATION' l_td.
          append l_td to lt_html. write <ls_job>-bw_destination to l_field. append l_field to lt_html. append '</td>' to lt_html.
        endif.

        append '</tr>' to lt_html.
      endloop.
      append '</table></p>' to lt_html.
    endif.
  endif.
  append '<p>Cumprimentos,</p>'(f01) to lt_html.
  append '<p>Equipa de suporte</p>'(f02) to lt_html.
  append '</body>' to lt_html.

  try.
* Create send request
      cl_send_request = cl_bcs=>create_persistent( ).

      data: lr_sender type ref to if_sender_bcs.
      lr_sender    = cl_cam_address_bcs=>create_internet_address( p_sender ).
      cl_send_request->set_sender( lr_sender ).
* Create new document with mailtitle and mailtextg
*      cl_document = cl_document_bcs=>create_document( i_type    = 'RAW' "#EC NOTEXT
*                                                      i_text    = t_mailtext
*                                                      i_subject = mail_title ).
      cl_document = cl_document_bcs=>create_document( i_type    = 'HTM' "'RAW' "#EC NOTEXT
                                                      i_text    = lt_html "t_mailtext
                                                      i_subject = mail_title ).

      if p_attach is not initial and gt_jobs is not initial.
* Add attachment to document
* since the new excelfiles have an 4-character extension .xlsx but the attachment-type only holds 3 charactes .xls,
* we have to specify the real filename via attachment header
* Use attachment_type xls to have SAP display attachment with the excel-icon
        attachment_subject  = l_save_file_name.
        concatenate '&SO_FILENAME=' attachment_subject into wa_attachment_header.
        append wa_attachment_header to t_attachment_header.
* Attachment
        sood_bytecount = bytecount.  " next method expects sood_bytecount instead of any positive integer *sigh*
        cl_document->add_attachment(  i_attachment_type    = 'XLS' "#EC NOTEXT
                                      i_attachment_subject = attachment_subject
                                      i_attachment_size    = sood_bytecount
                                      i_att_content_hex    = t_rawdata
                                      i_attachment_header  = t_attachment_header ).
      endif.

* add document to send request
      cl_send_request->set_document( cl_document ).

* add recipient(s)
      if p_email is not initial.
        send_to = p_email.
        cl_recipient = cl_cam_address_bcs=>create_internet_address( send_to ).
        cl_send_request->add_recipient( cl_recipient ).
      endif.
*      IF p_email2 IS NOT INITIAL.
*        send_to = p_email2.
*        cl_recipient = cl_cam_address_bcs=>create_internet_address( send_to ).
*        cl_send_request->add_recipient( cl_recipient ).
*      ENDIF.

* Und abschicken
      sent = cl_send_request->send( i_with_error_screen = 'X' ).

      commit work.

      if sent is initial.
        message i500(sbcoms) with p_email.
      else.
        message s022(so).
        message 'Document ready to be sent - Check SOST or SCOT' type 'I'.
      endif.

    catch cx_bcs into bcs_exception.
      errortext = bcs_exception->if_message~get_text( ).
      message errortext type 'I'.

  endtry.


endform.                    " F_SEND_EMAIL
*&---------------------------------------------------------------------*
*&      Form  f_create_list_in_excel
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->PO_EXCEL   text
*----------------------------------------------------------------------*
form f_create_list_in_excel  using po_excel type ref to zcl_excel.
  field-symbols <ls_job> type t_job.

  data: ls_styles type t_styles.
  data: lo_worksheet type ref to zcl_excel_worksheet.
  data: l_row type i.
  data: l_col type i.
*  data: lr_column_dimension type ref to zcl_excel_worksheet_columndime.
  data: l_style type zexcel_cell_style.
*  DATA: lt_autofilter TYPE zexcel_t_autofilter_values.
*  FIELD-SYMBOLS <ls_autofilter> TYPE LINE OF zexcel_t_autofilter_values.

  create object po_excel.
  perform f_create_styles using po_excel changing ls_styles.

  lo_worksheet = po_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Listagem de erros de job'(t01) ).

*header
  l_col = 1.
  lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Job'(c01) ip_style = ls_styles-header_up ). add 1 to l_col.
  lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Execução'(c02) ip_style = ls_styles-header_up ). add 1 to l_col.
  lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Execução'(c03) ip_style = ls_styles-header_up ). add 1 to l_col.
  lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Escalonador'(c04) ip_style = ls_styles-header_up ). add 1 to l_col.
  lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Status'(c05) ip_style = ls_styles-header_up ). add 1 to l_col.
  lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Log'(c06) ip_style = ls_styles-header_up ). add 1 to l_col.
  lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Log'(c07) ip_style = ls_styles-header_up ). add 1 to l_col.
  lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'ID'(c08) ip_style = ls_styles-header_up ). add 1 to l_col.
  lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Nº'(c09) ip_style = ls_styles-header_up ). add 1 to l_col.
  lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Tipo'(c10) ip_style = ls_styles-header_up ). add 1 to l_col.
  lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Erro'(c11) ip_style = ls_styles-header_up ). add 1 to l_col.
  lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Chave'(c12) ip_style = ls_styles-header_up ). add 1 to l_col.

  if p_steps is not initial.
    lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Cadeia'(c13) ip_style = ls_styles-header_up ). add 1 to l_col.
    lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Etapa'(c14) ip_style = ls_styles-header_up ). add 1 to l_col.
    lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Programa Cadeia'(c15) ip_style = ls_styles-header_up ). add 1 to l_col.
    lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Variante Cadeia'(c16) ip_style = ls_styles-header_up ). add 1 to l_col.
    lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = 'Destino Cadeia'(c17) ip_style = ls_styles-header_up ). add 1 to l_col.
  endif.

  l_row = 1.
  loop at gt_jobs assigning <ls_job>.
    add 1 to l_row.
    l_col = 1.

    perform f_get_style using <ls_job> 'JOBNAME' l_style ls_styles.
    lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-jobname ip_style = l_style ). add 1 to l_col.
    perform f_get_style using <ls_job> 'SDLSTRTDT' l_style ls_styles.
    lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-sdlstrtdt ip_style = l_style ). add 1 to l_col.
    perform f_get_style using <ls_job> 'SDLSTRTTM' l_style ls_styles.
    lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-sdlstrttm ip_style = l_style ). add 1 to l_col.
    perform f_get_style using <ls_job> 'SDLUNAME' l_style ls_styles.
    lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-sdluname ip_style = l_style ). add 1 to l_col.
    perform f_get_style using <ls_job> 'STATUS' l_style ls_styles.
    lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-status ip_style = l_style ). add 1 to l_col.
    perform f_get_style using <ls_job> 'ENTERDATE' l_style ls_styles.
    lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-enterdate ip_style = l_style ). add 1 to l_col.
    perform f_get_style using <ls_job> 'ENTERTIME' l_style ls_styles.
    lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-entertime ip_style = l_style ). add 1 to l_col.
    perform f_get_style using <ls_job> 'MSGID' l_style ls_styles.
    lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-msgid ip_style = l_style ). add 1 to l_col.
    perform f_get_style using <ls_job> 'MSGNO' l_style ls_styles.
    lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-msgno ip_style = l_style ). add 1 to l_col.
    perform f_get_style using <ls_job> 'MSGTYPE' l_style ls_styles.
    lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-msgtype ip_style = l_style ). add 1 to l_col.
*    APPEND INITIAL LINE TO lt_autofilter ASSIGNING <ls_autofilter>.
*    <ls_autofilter>-column = l_col.
*    <ls_autofilter>-value = 'A'.
*    APPEND INITIAL LINE TO lt_autofilter ASSIGNING <ls_autofilter>.
*    <ls_autofilter>-column = l_col.
*    <ls_autofilter>-value = 'E'.
*    APPEND INITIAL LINE TO lt_autofilter ASSIGNING <ls_autofilter>.
*    <ls_autofilter>-column = l_col.
*    <ls_autofilter>-value = 'I'.
*    APPEND INITIAL LINE TO lt_autofilter ASSIGNING <ls_autofilter>.
*    <ls_autofilter>-column = l_col.
*    <ls_autofilter>-value = 'S'.
*    APPEND INITIAL LINE TO lt_autofilter ASSIGNING <ls_autofilter>.
*    <ls_autofilter>-column = l_col.
*    <ls_autofilter>-value = 'W'.
*    APPEND INITIAL LINE TO lt_autofilter ASSIGNING <ls_autofilter>.
*    <ls_autofilter>-column = l_col.
*    <ls_autofilter>-value = 'X'.
**    lo_worksheet->get_cell( EXPORTING
**                               ip_column    = l_col
**                               ip_row       = l_row
**                            IMPORTING
**                               ep_value     = <ls_autofilter>-value ).
    perform f_get_style using <ls_job> 'TEXT' l_style ls_styles.
    lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-text ip_style = l_style ). add 1 to l_col.
    perform f_get_style using <ls_job> 'ERRKEY' l_style ls_styles.
    lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-errkey ip_style = l_style ). add 1 to l_col.

    if p_steps is not initial.
      perform f_get_style using <ls_job> 'BW_CHAIN' l_style ls_styles.
      lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-bw_chain ip_style = l_style ). add 1 to l_col.
      perform f_get_style using <ls_job> 'BW_STEP' l_style ls_styles.
      lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-bw_step ip_style = l_style ). add 1 to l_col.
      perform f_get_style using <ls_job> 'BW_PROGRAM' l_style ls_styles.
      lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-bw_program ip_style = l_style ). add 1 to l_col.
      perform f_get_style using <ls_job> 'BW_VARIANT' l_style ls_styles.
      lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-bw_variant ip_style = l_style ). add 1 to l_col.
      perform f_get_style using <ls_job> 'BW_DESTINATION' l_style ls_styles.
      lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <ls_job>-bw_destination ip_style = l_style ). add 1 to l_col.
    endif.

  endloop.

  clear l_col.
  do 11 times.
    add 1 to l_col.
    lo_worksheet->set_column_width( ip_column = l_col ).
  enddo.

*  DATA: lo_autofilter TYPE REF TO zcl_excel_autofilter,
*        ls_area TYPE zexcel_s_autofilter_area.
*  lo_autofilter = po_excel->add_new_autofilter( lo_worksheet ).
*  ls_area-row_start = 1.
*  ls_area-col_start = 1.
*  ls_area-row_end = lo_worksheet->get_highest_row( ).
*  ls_area-col_end = lo_worksheet->get_highest_column( ).
*  lo_autofilter->set_filter_area( ls_area ).
*  lo_autofilter->set_values( lt_autofilter ).
**  lo_autofilter->set_value( i_column = 9 i_value = 'E' ).

endform.                    "f_create_list_in_excel
*&---------------------------------------------------------------------*
*&      Form  f_create_styles
*&---------------------------------------------------------------------*
form f_create_styles  using    po_excel  type ref to zcl_excel
                      changing ps_styles type t_styles.
  data: lo_worksheet            type ref to zcl_excel_worksheet,
        lo_style_header_up      type ref to zcl_excel_style,
        lo_style_header_edit    type ref to zcl_excel_style,
        lo_style_header_down    type ref to zcl_excel_style,
        lo_style_normal         type ref to zcl_excel_style,
        lo_style_edit           type ref to zcl_excel_style,
        lo_style_red            type ref to zcl_excel_style,
        lo_style_light_red      type ref to zcl_excel_style,
        lo_style_green          type ref to zcl_excel_style,
        lo_style_yellow         type ref to zcl_excel_style,
        lo_border_dark          type ref to zcl_excel_style_border.

  " Create border object
  create object lo_border_dark.
  lo_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
  lo_border_dark->border_style = zcl_excel_style_border=>c_border_thin.

  "Create style top header
  lo_style_header_up                         = po_excel->add_new_style( ).
  lo_style_header_up->borders->allborders    = lo_border_dark.
  lo_style_header_up->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_header_up->fill->fgcolor-rgb  = '0F243E'. "top zahara blue
  lo_style_header_up->font->bold   = abap_true.
  lo_style_header_up->font->color-rgb  = zcl_excel_style_color=>c_white.
  lo_style_header_up->font->name         = zcl_excel_style_font=>c_name_calibri.
  lo_style_header_up->font->size = 10.
*  lo_style_header_up->font->scheme       = zcl_excel_style_font=>c_scheme_major.
*  lo_style_header_up->protection->locked = zcl_excel_style_protection=>c_protection_locked.
  ps_styles-header_up                    = lo_style_header_up->get_guid( ).


  "Create style top header edit column
  lo_style_header_edit                         = po_excel->add_new_style( ).
  lo_style_header_edit->borders->allborders    = lo_border_dark.
  lo_style_header_edit->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_header_edit->fill->fgcolor-rgb  = '0F243E'. "top zahara blue
  lo_style_header_edit->font->bold        = abap_true.
  lo_style_header_edit->font->color-rgb  = zcl_excel_style_color=>c_red.
  lo_style_header_edit->font->name         = zcl_excel_style_font=>c_name_calibri.
  lo_style_header_edit->font->size = 10.
*  lo_style_header_edit->font->scheme       = zcl_excel_style_font=>c_scheme_major.
*  lo_style_header_edit->protection->locked = zcl_excel_style_protection=>c_protection_locked.
  ps_styles-header_edit                    = lo_style_header_edit->get_guid( ).

  "Create style second line header
  lo_style_header_down                         = po_excel->add_new_style( ).
  lo_style_header_down->borders->allborders    = lo_border_dark.
  lo_style_header_down->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_header_down->fill->fgcolor-rgb  = 'ADD1A5'. "small header zahara green
  lo_style_header_down->font->name         = zcl_excel_style_font=>c_name_calibri.
  lo_style_header_down->font->size = 10.
*  lo_style_header_down->protection->locked = zcl_excel_style_protection=>c_protection_locked.
  ps_styles-header_down                    = lo_style_header_down->get_guid( ).

  "Create style table cells
  lo_style_normal                         = po_excel->add_new_style( ).
  lo_style_normal->borders->allborders    = lo_border_dark.
  lo_style_normal->font->name             = zcl_excel_style_font=>c_name_calibri.
  lo_style_normal->font->size             = 8.
*  lo_style_normal->protection->locked = zcl_excel_style_protection=>c_protection_locked.
  ps_styles-normal                         = lo_style_normal->get_guid( ).

  lo_style_red                         = po_excel->add_new_style( ).
  lo_style_red->borders->allborders    = lo_border_dark.
  lo_style_red->font->name             = zcl_excel_style_font=>c_name_calibri.
  lo_style_red->font->size             = 8.
  lo_style_red->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_red->fill->fgcolor-rgb  = 'FF3333'. "red
  ps_styles-red                         = lo_style_red->get_guid( ).

  lo_style_light_red                         = po_excel->add_new_style( ).
  lo_style_light_red->borders->allborders    = lo_border_dark.
  lo_style_light_red->font->name             = zcl_excel_style_font=>c_name_calibri.
  lo_style_light_red->font->size             = 8.
  lo_style_light_red->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_light_red->fill->fgcolor-rgb  = 'FFCCCC'. "light_red
  ps_styles-light_red                         = lo_style_light_red->get_guid( ).

  lo_style_green                         = po_excel->add_new_style( ).
  lo_style_green->borders->allborders    = lo_border_dark.
  lo_style_green->font->name             = zcl_excel_style_font=>c_name_calibri.
  lo_style_green->font->size             = 8.
  lo_style_green->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_green->fill->fgcolor-rgb  = '00CC44'. "green
  ps_styles-green                         = lo_style_green->get_guid( ).

  lo_style_yellow                         = po_excel->add_new_style( ).
  lo_style_yellow->borders->allborders    = lo_border_dark.
  lo_style_yellow->font->name             = zcl_excel_style_font=>c_name_calibri.
  lo_style_yellow->font->size             = 8.
  lo_style_yellow->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_yellow->fill->fgcolor-rgb  = 'FFFF00'. "yellow
  ps_styles-yellow                         = lo_style_yellow->get_guid( ).

  "Create style edit cells
  lo_style_edit                         = po_excel->add_new_style( ).
  lo_style_edit->borders->allborders    = lo_border_dark.
  lo_style_edit->font->name             = zcl_excel_style_font=>c_name_calibri.
  lo_style_edit->font->size             = 8.
  lo_style_edit->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_edit->fill->fgcolor-rgb  = 'FFFF66'. "top zahara blue
  lo_style_edit->number_format->format_code = zcl_excel_style_number_format=>c_format_text."c_format_number.
  lo_style_edit->protection->locked = zcl_excel_style_protection=>c_protection_unlocked.
  ps_styles-edit                        = lo_style_edit->get_guid( ).

endform.                    " F_CREATE_STYLES
*&---------------------------------------------------------------------*
*&      Form  F_GET_STYLE
*&---------------------------------------------------------------------*
form f_get_style  using    ps_job type t_job
                           value(p_field) type string
                           p_style type zexcel_cell_style
                           ps_styles type t_styles.
  field-symbols: <ls_color> like line of ps_job-colortab.
*      <ls_color>-color-col = 3.  "amarelo
*      <ls_color>-color-col = 5.  "verde
*      <ls_color>-color-col = 6.  "vermelho
*      <ls_color>-color-col = 7.  "vermelho leve
*      <ls_color>-color-inv = '1'. "inverte se = 0
*      <ls_color>-color-int = '1'. "intensifica

  read table ps_job-colortab assigning <ls_color> with key fname = p_field.
  if sy-subrc ne 0.
    p_style = ps_styles-normal.
  else.
    case <ls_color>-color-col.
      when 3.
        p_style = ps_styles-yellow.
      when 5.
        p_style = ps_styles-green.
      when 6.
        p_style = ps_styles-red.
      when 7.
        p_style = ps_styles-light_red.
    endcase.
  endif.
endform.                    " F_GET_STYLE
*&---------------------------------------------------------------------*
*&      Form  F_GET_TD_CLASS
*&---------------------------------------------------------------------*
form f_get_td_class  using    ps_job type t_job
                              value(p_field) type string
                              p_style.
  field-symbols: <ls_color> like line of ps_job-colortab.
  read table ps_job-colortab assigning <ls_color> with key fname = p_field.
  if sy-subrc ne 0.
    p_style = '<TD CLASS="normal">'.
  else.
    case <ls_color>-color-col.
      when 3.
        p_style = '<TD CLASS="yellow">'.
      when 5.
        p_style = '<TD CLASS="green">'.
      when 6.
        p_style = '<TD CLASS="red">'.
      when 7.
        p_style = '<TD CLASS="light_red">'.
    endcase.
  endif.

endform.                    "f_get_td_class
*&---------------------------------------------------------------------*
*&      Form  F_SET_COLORS
*&---------------------------------------------------------------------*
form f_set_colors .
  field-symbols <ls_color> type lvc_s_scol.
  field-symbols: <ls_job> type t_job.
  data: l_color type i.

*      <ls_color>-color-col = 3.  "amarelo
*      <ls_color>-color-col = 5.  "verde
*      <ls_color>-color-col = 6.  "vermelho
*      <ls_color>-color-col = 7.  "vermelho leve
*      <ls_color>-color-inv = '1'. "inverte se = 0
*      <ls_color>-color-int = '1'. "intensifica
  loop at gt_jobs assigning <ls_job>.
    clear l_color.
    if ( s_errkyr[] is not initial and <ls_job>-errkey in s_errkyr ) or <ls_job>-msgtype eq 'A' or <ls_job>-msgtype eq 'X'.
      l_color = 6. "vermelho
    elseif <ls_job>-msgtype eq 'E'.
      l_color = 7. "vermelho leve
    elseif <ls_job>-msgtype eq 'W'.
      l_color = 3. "amarelo
    elseif <ls_job>-msgtype eq 'S'.
      l_color = 5.  "verde
    endif.
    if l_color is not initial.
      append initial line to <ls_job>-colortab assigning <ls_color>.
      <ls_color>-fname = 'JOBNAME'.
      <ls_color>-color-col = l_color.
      append initial line to <ls_job>-colortab assigning <ls_color>.
      <ls_color>-fname = 'MSGID'.
      <ls_color>-color-col = l_color.
      append initial line to <ls_job>-colortab assigning <ls_color>.
      <ls_color>-fname = 'MSGNO'.
      <ls_color>-color-col = l_color.
      append initial line to <ls_job>-colortab assigning <ls_color>.
      <ls_color>-fname = 'MSGTYPE'.
      <ls_color>-color-col = l_color.
      append initial line to <ls_job>-colortab assigning <ls_color>.
      <ls_color>-fname = 'TEXT'.
      <ls_color>-color-col = l_color.
    endif.

  endloop.
endform.                    " F_SET_COLORS
*&---------------------------------------------------------------------*
*&      Form  F_SET_R_REPID
*&---------------------------------------------------------------------*
form f_set_r_repid .
  field-symbols: <lr_repid> like line of gr_repid.

  clear gr_repid[].
  loop at s_repid.
    append initial line to gr_repid assigning <lr_repid>.
    <lr_repid>-sign = s_repid-sign.
    <lr_repid>-option = s_repid-option.
    <lr_repid>-low = s_repid-low.
    <lr_repid>-high = s_repid-high.
  endloop.
endform.                    " F_SET_R_REPID
