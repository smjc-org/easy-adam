/*
 * Macro Name:    ads_import_excel
 * Macro Purpose: ��ȡ Excel ���ݣ����� SAS ���ݼ�
 * Author:        wtwang
 * Version Date:  2025-04-01
*/

%macro ads_import_excel(file,
                        outdata,
                        sheet_name,
                        range_attr                   = #null,
                        range_data                   = #null,
                        all_chars                    = true,
                        clear_format                 = true,
                        clear_informat               = true,
                        ignore_empty_line            = true,
                        warning_var_name_empty       = true,
                        warning_var_name_not_meet_v7 = true,
                        warning_var_name_len_gt_8    = true,
                        debug                        = false) / parmbuff;
    /*  file:                         Excel �ļ�·��
     *  outdata:                      SAS ���ݼ�����
     *  sheet_name:                   ����������
     *  range_attr:                   �����������ͱ�ǩ����� 2*C ��Ԫ���������е�һ�б����Ǳ�ǩ���ڶ��б����Ǳ�����
     *  range_data:                   �����������ݵ� R*C ��Ԫ������
     *  all_chars:                    �Ƿ����б�����Ϊ�ַ��ͱ���
     *  clear_format:                 �Ƿ���������󶨵������ʽ
     *  clear_informat:               �Ƿ���������󶨵������ʽ
     *  ignore_empty_line:            �Ƿ���Կ���
     *  warning_var_name_empty:       ��������Ϊ��ʱ���Ƿ����������Ϣ
     *  warning_var_name_not_meet_v7: ������������ VALIDVARNAME=V7 �µķǷ��ַ�ʱ���Ƿ����������Ϣ
     *  warning_var_name_len_gt_8:    �����������ȳ��� 8 ʱ���Ƿ����������Ϣ
     *  debug:                        ����ģʽ
    */

    /*ͳһ������Сд*/
    %let file                         = %sysfunc(strip(%bquote(&file)));
    %let outdata                      = %sysfunc(strip(%bquote(&outdata)));
    %let sheet_name                   = %sysfunc(strip(%bquote(&sheet_name)));
    %let range_attr                   = %upcase(%sysfunc(strip(%bquote(&range_attr))));
    %let range_data                   = %upcase(%sysfunc(strip(%bquote(&range_data))));
    %let all_chars                    = %upcase(%sysfunc(strip(%bquote(&all_chars))));
    %let clear_format                 = %upcase(%sysfunc(strip(%bquote(&clear_format))));
    %let clear_informat               = %upcase(%sysfunc(strip(%bquote(&clear_informat))));
    %let ignore_empty_line            = %upcase(%sysfunc(strip(%bquote(&ignore_empty_line))));
    %let warning_var_name_empty       = %upcase(%sysfunc(strip(%bquote(&warning_var_name_empty))));
    %let warning_var_name_not_meet_v7 = %upcase(%sysfunc(strip(%bquote(&warning_var_name_not_meet_v7))));
    %let warning_var_name_len_gt_8    = %upcase(%sysfunc(strip(%bquote(&warning_var_name_len_gt_8))));
    %let debug                        = %upcase(%sysfunc(strip(%bquote(&debug))));

    /*��ȡ Excel �ļ�*/
    %if &all_chars = TRUE %then %do;
        %let EFI_ALLCHARS = YES;
    %end;

    %if (&range_attr = #NULL and &range_data ^= #NULL) or (&range_attr ^= #NULL and &range_data = #NULL) %then %do;
        %put ERROR: ���� RANGE_ATTR �� RANGE_DATA ����ͬʱΪ #NULL ����ͬʱ��Ϊ #NULL��;
        %goto exit;
    %end;
    %else %if &range_attr = #NULL and &range_data = #NULL %then %do;
        proc import file = "&file" out = tmp_excel dbms = xlsx replace;
            sheet = "&sheet_name";
            getnames = no;
        run;

        data tmp_excel_attr;
            set tmp_excel(firstobs = 1 obs = 2);
        run;

        data tmp_excel_data;
            set tmp_excel(firstobs = 3);
        run;
    %end;
    %else %do;
        proc import file = "&file" out = tmp_excel_attr dbms = xlsx replace;
            sheet = "&sheet_name";
            range = "&range_attr";
            getnames = no;
        run;

        proc import file = "&file" out = tmp_excel_data dbms = xlsx replace;
            sheet = "&sheet_name";
            range = "&range_data";
            getnames = no;
        run;
    %end;

    /*ת�� tmp_excel_attr*/
    proc transpose data = tmp_excel_attr out = tmp_excel_attr_trans;
        var _all_;
    run;

    /*��ȡ�������ƺͱ�����ǩ*/
    proc sql noprint;
        select _NAME_       into :var_old_name_1-  from tmp_excel_attr_trans;
        select kstrip(COL2) into :var_new_name_1-  from tmp_excel_attr_trans;
        select kstrip(COL1) into :var_new_label_1- from tmp_excel_attr_trans;
    quit;
    %let var_n = &sqlobs;

    /*���������Ƿ���Ϲ淶*/
    %do i = 1 %to &var_n;
        %let IS_VALID_VAR_&i = TRUE;

        /*���������Ƿ�Ϊ��*/
        %if &warning_var_name_empty = TRUE %then %do;
            %if %superq(var_new_name_&i) = %bquote() %then %do;
                %let IS_VALID_VAR_&i = FALSE;
                %put WARNING: �� %superq(var_old_name_&i) �ı�����Ϊ�գ�;
            %end;
        %end;

        /*���������Ƿ��� VALIDVARNAME=V7 �ºϷ�*/
        %if &warning_var_name_not_meet_v7 = TRUE %then %do;
            %if %sysfunc(notname(%superq(var_new_name_&i))) > 0 %then %do;
                %let IS_VALID_VAR_&i = FALSE;
                %put WARNING: �� %superq(var_old_name_&i) �ı����� %superq(var_new_name_&i) �� VALIDVARNAME=V7 �²��Ϸ���;
            %end;
        %end;

        /*���������ĳ����Ƿ񳬹� 8*/
        %if &warning_var_name_len_gt_8 = TRUE %then %do;
            %if %length(%superq(var_new_name_&i)) > 8 %then %do;
                %put WARNING: �� %superq(var_old_name_&i) �ı����� %superq(var_new_name_&i) �ĳ��ȳ��� 8 ��;
            %end;
        %end;

        %if &&IS_VALID_VAR_&i = FALSE %then %do;
            %let var_new_name_&i = &&var_old_name_&i;
        %end;
    %end;

    /*�޸ı������ƺͱ�����ǩ*/
    proc sql noprint;
        create table tmp_excel_data_renamed as
            select
                %do i = 1 %to &var_n;
                    &&var_old_name_&i as &&var_new_name_&i label = %unquote(%str(%')%superq(var_new_label_&i)%str(%'))
                    %if &i < &var_n %then %do; %bquote(,) %end;
                %end;
            from tmp_excel_data;
    quit;

    /*�Ƴ�����Ŀ���*/
    %if &ignore_empty_line = TRUE %then %do;
        data tmp_excel_data_renamed;
            set tmp_excel_data_renamed;
            if cmiss(%do i = 1 %to &var_n;
                         &&var_new_name_&i
                         %if &i < &var_n %then %do; %bquote(,) %end;
                     %end;) = &var_n then delete;
        run;
    %end;

    /*������б����� format �� informat*/
    proc datasets library = work noprint nowarn;
        modify tmp_excel_data_renamed;
        %if &clear_format = TRUE %then %do;
            attrib _all_ format=;
        %end;
        %if &clear_informat = TRUE %then %do;
            attrib _all_ informat=;
        %end;
    run;
    quit;

    /*����������ݼ�*/
    data &outdata;
        set tmp_excel_data_renamed;
    run;

    /*ɾ���м����ݼ�*/
    %if %bquote(&debug) = %upcase(false) %then %do;
        proc datasets library = work nowarn noprint;
            delete tmp_excel
                   tmp_excel_attr
                   tmp_excel_data
                   tmp_excel_attr_trans
                   tmp_excel_data_renamed
                   ;
        quit;
    %end;

    %exit:
    %let EFI_ALLCHARS = NO;
    %put NOTE: ����� ads_import_excel �ѽ������У�;
%mend;
