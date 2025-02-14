/*
 * Macro Name:    ads_import_excel
 * Macro Purpose: 读取 Excel 数据，创建 SAS 数据集
 * Author:        wtwang
 * Version Date:  2025-02-05 0.1.0
*/

%macro ads_import_excel(file,
                        outdata,
                        sheet_name,
                        range_attr                   = #null,
                        range_data                   = #null,
                        convert_to_char              = true,
                        warning_var_name_empty       = true,
                        warning_var_name_not_meet_v7 = true,
                        warning_var_name_len_gt_8    = true,
                        debug                        = false) / parmbuff;
    /*  file:                         Excel 文件路径
     *  outdata:                      SAS 数据集名称
     *  sheet_name:                   工作表名称
     *  range_attr:                   包含变量名和标签定义的 2*C 单元格区域，其中第一行必须是标签，第二行必须是变量名
     *  range_data:                   包含数据内容的 R*C 单元格区域
     *  convert_to_char:              是否强制将所有变量转为字符型变量
     *  warning_var_name_empty:       当变量名为空时，是否输出警告信息
     *  warning_var_name_not_meet_v7: 当变量名包含 VALIDVARNAME=V7 下的非法字符时，是否输出警告信息
     *  warning_var_name_len_gt_8:    当变量名长度超过 8 时，是否输出警告信息
     *  debug:                        调试模式
    */

    /*统一参数大小写*/
    %let file                         = %sysfunc(strip(%bquote(&file)));
    %let outdata                      = %sysfunc(strip(%bquote(&outdata)));
    %let sheet_name                   = %sysfunc(strip(%bquote(&sheet_name)));
    %let range_attr                   = %upcase(%sysfunc(strip(%bquote(&range_attr))));
    %let range_data                   = %upcase(%sysfunc(strip(%bquote(&range_data))));
    %let convert_to_char              = %upcase(%sysfunc(strip(%bquote(&convert_to_char))));
    %let warning_var_name_empty       = %upcase(%sysfunc(strip(%bquote(&warning_var_name_empty))));
    %let warning_var_name_not_meet_v7 = %upcase(%sysfunc(strip(%bquote(&warning_var_name_not_meet_v7))));
    %let warning_var_name_len_gt_8    = %upcase(%sysfunc(strip(%bquote(&warning_var_name_len_gt_8))));
    %let debug                        = %upcase(%sysfunc(strip(%bquote(&debug))));

    /*读取 Excel 文件*/
    %if (&range_attr = #NULL and &range_data ^= #NULL) or (&range_attr ^= #NULL and &range_data = #NULL) %then %do;
        %put ERROR: 参数 RANGE_ATTR 和 RANGE_DATA 必须同时为 #NULL 或者同时不为 #NULL！;
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

    /*转置 tmp_excel_attr*/
    proc transpose data = tmp_excel_attr out = tmp_excel_attr_trans;
        var _all_;
    run;

    /*获取变量名称和变量标签*/
    proc sql noprint;
        select _NAME_       into :var_old_name_1-  from tmp_excel_attr_trans;
        select kstrip(COL2) into :var_new_name_1-  from tmp_excel_attr_trans;
        select kstrip(COL1) into :var_new_label_1- from tmp_excel_attr_trans;
    quit;
    %let var_n = &sqlobs;

    /*检查变量名是否符合规范*/
    %do i = 1 %to &var_n;
        %let IS_VALID_VAR_&i = TRUE;

        /*检查变量名是否为空*/
        %if &warning_var_name_empty = TRUE %then %do;
            %if %superq(var_new_name_&i) = %bquote() %then %do;
                %let IS_VALID_VAR_&i = FALSE;
                %put WARNING: 列 %superq(var_old_name_&i) 的变量名为空！;
            %end;
        %end;

        /*检查变量名是否在 VALIDVARNAME=V7 下合法*/
        %if &warning_var_name_not_meet_v7 = TRUE %then %do;
            %if %sysfunc(notname(%superq(var_new_name_&i))) > 0 %then %do;
                %let IS_VALID_VAR_&i = FALSE;
                %put WARNING: 列 %superq(var_old_name_&i) 的变量名 %superq(var_new_name_&i) 在 VALIDVARNAME=V7 下不合法！;
            %end;
        %end;

        /*检查变量名的长度是否超过 8*/
        %if &warning_var_name_len_gt_8 = TRUE %then %do;
            %if %length(%superq(var_new_name_&i)) > 8 %then %do;
                %put WARNING: 列 %superq(var_old_name_&i) 的变量名 %superq(var_new_name_&i) 的长度超过 8 ！;
            %end;
        %end;

        %if &&IS_VALID_VAR_&i = FALSE %then %do;
            %let var_new_name_&i = &&var_old_name_&i;
        %end;
    %end;

    /*修改变量名称和变量标签*/
    proc sql noprint;
        create table tmp_excel_data_processed as
            select
                %do i = 1 %to &var_n;
                    &&var_old_name_&i as &&var_new_name_&i label = %unquote(%str(%')%superq(var_new_label_&i)%str(%'))
                    %if &i < &var_n %then %do; %bquote(,) %end;
                %end;
            from tmp_excel_data;
    quit;

    /*强制将所有变量转为字符型变量*/
    %if &convert_to_char = TRUE %then %do;
        proc sql noprint;
            select name   into :name_1-   from dictionary.columns where libname = "WORK" and memname = "TMP_EXCEL_DATA_PROCESSED";
            select type   into :type_1-   from dictionary.columns where libname = "WORK" and memname = "TMP_EXCEL_DATA_PROCESSED";
            select format into :format_1- from dictionary.columns where libname = "WORK" and memname = "TMP_EXCEL_DATA_PROCESSED";
            %let var_n = &sqlobs;

            create table tmp_excel_data_converted as
                select
                    %do i = 1 %to &var_n;
                        put(&&name_&i, &&format_&i -L) as &&name_&i
                        %if &i < &var_n %then %do; %bquote(,) %end;
                    %end;
                from tmp_excel_data_processed;
        quit;
    %end;

    /*创建输出数据集*/
    data &outdata;
        set %if &convert_to_char = TRUE %then %do;
                tmp_excel_data_converted
            %end;
            %else %do;
                tmp_excel_data_processed
            %end;
            ;
    run;

    /*删除中间数据集*/
    %if %bquote(&debug) = %upcase(false) %then %do;
        proc datasets library = work nowarn noprint;
            delete tmp_excel
                   tmp_excel_attr
                   tmp_excel_data
                   tmp_excel_attr_trans
                   tmp_excel_data_processed
                   tmp_excel_data_converted
                   ;
        quit;
    %end;

    %exit:
    %put NOTE: 宏程序 ads_import_excel 已结束运行！;
%mend;
