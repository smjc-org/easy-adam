/*
 * Macro Name:    ads_import_excel
 * Macro Purpose: 读取 Excel 数据，创建 SAS 数据集
 * Author:        wtwang
 * Version Date:  2025-08-04
*/

%macro ads_import_excel(file,
                        outdata,
                        sheet_name,
                        dbms                         = #auto,
                        range_attr                   = #null,
                        range_attr_row_index         = %str(2, 1),
                        range_data                   = #null,
                        sort_by                      = #null,
                        all_chars                    = false,
                        varlist_fixed_char           = #null,
                        varlist_fixed_num            = #null,
                        clear_format                 = true,
                        clear_informat               = true,
                        ignore_empty_line            = true,
                        warning_var_name_empty       = true,
                        warning_var_name_not_meet_v7 = true,
                        warning_var_name_len_gt_8    = true,
                        debug                        = false) / parmbuff;
    /*  file:                         Excel 文件路径
     *  outdata:                      SAS 数据集名称
     *  sheet_name:                   工作表名称
     *  range_attr:                   包含变量名和标签定义的单元格区域，其中必须包含一行标签和一行变量名
     *  range_attr_row_index:         单元格区域 range_attr 内代表变量名和标签的行索引（行索引从 range_attr 指定的单元格区域的第一行开始计数，起始行号为 1，步进 1）
     *  range_data:                   包含数据内容的单元格区域
     *  sort_by:                      用于对输出数据集排序的变量（可指定排序方向）
     *  all_chars:                    是否将所有变量视为字符型变量
     *  varlist_fixed_char:           一个变量名列表，其中的变量均被视为字符型变量
     *  varlist_fixed_num:            一个变量名列表，其中的变量均被视为数值型变量
     *  clear_format:                 是否清除变量绑定的输出格式
     *  clear_informat:               是否清除变量绑定的输入格式
     *  ignore_empty_line:            是否忽略空行
     *  warning_var_name_empty:       当变量名为空时，是否输出警告信息
     *  warning_var_name_not_meet_v7: 当变量名包含 VALIDVARNAME=V7 下的非法字符时，是否输出警告信息
     *  warning_var_name_len_gt_8:    当变量名长度超过 8 时，是否输出警告信息
     *  debug:                        调试模式
    */

    /*统一参数大小写*/
    %let file                         = %sysfunc(strip(%bquote(&file)));
    %let outdata                      = %sysfunc(strip(%bquote(&outdata)));
    %let dbms                         = %upcase(%sysfunc(strip(%bquote(&dbms))));
    %let range_attr                   = %upcase(%sysfunc(strip(%bquote(&range_attr))));
    %let range_attr_row_index         = %upcase(%sysfunc(strip(%bquote(&range_attr_row_index))));
    %let range_data                   = %upcase(%sysfunc(strip(%bquote(&range_data))));
    %let sort_by                      = %upcase(%sysfunc(strip(%bquote(&sort_by))));
    %let all_chars                    = %upcase(%sysfunc(strip(%bquote(&all_chars))));
    %let varlist_fixed_char           = %upcase(%sysfunc(strip(%bquote(&varlist_fixed_char))));
    %let varlist_fixed_num            = %upcase(%sysfunc(strip(%bquote(&varlist_fixed_num))));
    %let clear_format                 = %upcase(%sysfunc(strip(%bquote(&clear_format))));
    %let clear_informat               = %upcase(%sysfunc(strip(%bquote(&clear_informat))));
    %let ignore_empty_line            = %upcase(%sysfunc(strip(%bquote(&ignore_empty_line))));
    %let warning_var_name_empty       = %upcase(%sysfunc(strip(%bquote(&warning_var_name_empty))));
    %let warning_var_name_not_meet_v7 = %upcase(%sysfunc(strip(%bquote(&warning_var_name_not_meet_v7))));
    %let warning_var_name_len_gt_8    = %upcase(%sysfunc(strip(%bquote(&warning_var_name_len_gt_8))));
    %let debug                        = %upcase(%sysfunc(strip(%bquote(&debug))));

    /*识别 DBMS*/
    %if &dbms = #AUTO %then %do;
        %if %length(&file) > 8 %then %do;
            %let path = &file;
        %end;
        %else %do;
            %if %sysfunc(fileref(&file)) = 0 %then %do;
                %let path = %sysfunc(pathname(&file));
            %end;
            %else %do;
                %let path = &file;
            %end;
        %end;

        %let path = %sysfunc(translate(&path, '/', '\'));
        %let filename = %scan(&path, -1, /);
        %let ext = %scan(&filename, -1, .);

        %if "&ext" = "&filename" %then %let ext = xlsx;

        %if &ext = xlsx or &ext = xlsm %then %do;
            %let dbms = xlsx;
        %end;
        %else %if &ext = xls %then %do;
            %let dbms = excel;
        %end;
        %else %do;
            %put ERROR: 不支持的数据格式 &ext ！;
            %goto exit;
        %end;
    %end;

    /*读取 Excel 文件*/
    %if (&range_attr = #NULL and &range_data ^= #NULL) or (&range_attr ^= #NULL and &range_data = #NULL) %then %do;
        %put ERROR: 参数 RANGE_ATTR 和 RANGE_DATA 必须同时为 #NULL 或者同时不为 #NULL！;
        %goto exit;
    %end;
    %else %if &range_attr = #NULL and &range_data = #NULL %then %do;
        proc import file = "&file" out = tmp_excel dbms = &dbms replace;
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
        proc import file = "&file" out = tmp_excel_attr dbms = &dbms replace;
            range = "&sheet_name$&range_attr";
            getnames = no;
        run;

        proc import file = "&file" out = tmp_excel_data dbms = &dbms replace;
            range = "&sheet_name$&range_data";
            getnames = no;
        run;
    %end;

    /*转置 tmp_excel_attr*/
    proc transpose data = tmp_excel_attr out = tmp_excel_attr_trans;
        var _all_;
    run;
    /*如果 range_attr 只有一行，则尝试将这一行视为变量名，若不可行，则将这一行视为变量标签，并使用默认的变量名（F1~Fn 或 A~Z）*/
    proc sql noprint;
        select * from dictionary.columns where libname = "WORK" and memname = "TMP_EXCEL_ATTR_TRANS" and name = "COL2";
    quit;
    %if &sqlobs = 0 %then %do;
        proc sql noprint;
            select ifc(sum(notname(strip(COL1))) = 0, "TRUE", "FALSE") into :ALL_COL1_VALID_FOR_V7_NAME from tmp_excel_attr_trans;
        quit;
        %if &ALL_COL1_VALID_FOR_V7_NAME = TRUE %then %do;
            data tmp_excel_attr_trans;
                set tmp_excel_attr_trans;
                COL2 = COL1;
            run;
        %end;
        %else %do;
            data tmp_excel_attr_trans;
                set tmp_excel_attr_trans;
                COL2 = _NAME_;
            run;
        %end;

        %let range_attr_row_index = %bquote(2, 1);
    %end;

    /*获取变量名称、变量标签、变量输出格式*/
    %let range_attr_row_var_name_index  = %scan(%bquote(&range_attr_row_index), 1, %bquote(, ));
    %let range_attr_row_var_label_index = %scan(%bquote(&range_attr_row_index), 2, %bquote(, ));
    proc sql noprint;
        select count(*) - 2 into :range_attr_row_index_max from dictionary.columns where libname = "WORK" and memname = "TMP_EXCEL_ATTR_TRANS";
    quit;
    %let IS_VALID_ROW_INDEX = TRUE;
    %if &range_attr_row_var_name_index > &range_attr_row_index_max %then %do;
        %put ERROR: 指定的变量名称的行索引 &range_attr_row_var_name_index 超出范围！;
        %let IS_VALID_ROW_INDEX = FALSE;
    %end;
    %if &range_attr_row_var_label_index > &range_attr_row_index_max %then %do;
        %put ERROR: 指定的变量标签的行索引 &range_attr_row_var_label_index 超出范围！;
        %let IS_VALID_ROW_INDEX = FALSE;
    %end;

    %if &IS_VALID_ROW_INDEX = FALSE %then %do;
        %goto exit;
    %end;

    proc sql noprint;
        select _NAME_                                     into :var_old_name_1-  from tmp_excel_attr_trans;
        select kstrip(COL&range_attr_row_var_name_index)  into :var_new_name_1-  from tmp_excel_attr_trans;
        select kstrip(COL&range_attr_row_var_label_index) into :var_new_label_1- from tmp_excel_attr_trans;
    quit;

    proc sql noprint;
        select kstrip(type)   into :var_raw_type_1-   from dictionary.columns where libname = "WORK" and memname = "TMP_EXCEL_DATA";
        select kstrip(format) into :var_raw_format_1- from dictionary.columns where libname = "WORK" and memname = "TMP_EXCEL_DATA";
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

    /*添加标识变量，记录已经被固定为字符型和数值型的变量*/
    %if %bquote(&varlist_fixed_char) ^= #NULL %then %do;
        %let var_fixed_char_n = %sysfunc(countw(%bquote(&varlist_fixed_char), %bquote(, )));
        %do i = 1 %to &var_fixed_char_n;
            %let var_fixed_char_&i = %scan(%bquote(&varlist_fixed_char), &i, %bquote(, ));
        %end;
    %end;

    %if %bquote(&varlist_fixed_num) ^= #NULL %then %do;
        %let var_fixed_num_n  = %sysfunc(countw(%bquote(&varlist_fixed_num), %bquote(, )));
        %do i = 1 %to &var_fixed_num_n;
            %let var_fixed_num_&i = %scan(%bquote(&varlist_fixed_num), &i, %bquote(, ));
        %end;
    %end;

    data tmp_excel_attr_trans;
        set tmp_excel_attr_trans;
        length fixed_char_flag fixed_num_flag $1;
        fixed_char_flag = "";
        fixed_num_flag = "";
        %if %bquote(&varlist_fixed_char) ^= #NULL or %bquote(&varlist_fixed_num) ^= #NULL %then %do;
            select (COL&range_attr_row_var_name_index);
                %if %bquote(&varlist_fixed_char) ^= #NULL %then %do;
                    when (%unquote(%do i = 1 %to &var_fixed_char_n;
                                       "&&var_fixed_char_&i"
                                       %if &i < &var_fixed_char_n %then %do; %bquote(,) %end;
                                   %end;)) fixed_char_flag = "Y";
                %end;
                %if %bquote(&varlist_fixed_num) ^= #NULL %then %do;
                    when (%unquote(%do i = 1 %to &var_fixed_num_n;
                                       "&&var_fixed_num_&i"
                                       %if &i < &var_fixed_num_n %then %do; %bquote(,) %end;
                                   %end;)) fixed_num_flag  = "Y";
                %end;
                otherwise;
            end;
        %end;
    run;

    proc sql noprint;
        select fixed_char_flag into :var_fixed_char_flag_1- from tmp_excel_attr_trans;
        select fixed_num_flag  into :var_fixed_num_flag_1-  from tmp_excel_attr_trans;
    quit;

    /*修改变量名称、变量标签、变量类型（可选）*/
    proc sql noprint;
        create table tmp_excel_data_renamed as
            select
                %if &all_chars = TRUE %then %do;
                    %do i = 1 %to &var_n;
                        put(&&var_old_name_&i, &&var_raw_format_&i -L) as &&var_new_name_&i label = %unquote(%str(%')%superq(var_new_label_&i)%str(%'))
                        %if &i < &var_n %then %do; %bquote(,) %end;
                    %end;
                %end;
                %else %do;
                    %do i = 1 %to &var_n;
                        %if &&var_fixed_char_flag_&i = Y %then %do;
                            %if &&var_raw_type_&i = char %then %do;
                                &&var_old_name_&i
                                %put WARNING: 列 %superq(var_old_name_&i) 的变量名 %superq(var_new_name_&i) 已经是字符型变量，无需转换！;
                            %end;
                            %else %do;
                                put(&&var_old_name_&i, &&var_raw_format_&i -L)
                            %end;
                        %end;
                        %else %if &&var_fixed_num_flag_&i = Y %then %do;
                            %if &&var_raw_type_&i = num %then %do;
                                &&var_old_name_&i
                                %put WARNING: 列 %superq(var_old_name_&i) 的变量名 %superq(var_new_name_&i) 已经是数值型变量，无需转换！;
                            %end;
                            %else %do;
                                input(&&var_old_name_&i, best.)
                            %end;
                        %end;
                        %else %do;
                            &&var_old_name_&i
                        %end;
                        as &&var_new_name_&i label = %unquote(%str(%')%superq(var_new_label_&i)%str(%'))
                        %if &i < &var_n %then %do; %bquote(,) %end;
                    %end;
                %end;
            from tmp_excel_data;
    quit;

    /*移除多余的空行*/
    %if &ignore_empty_line = TRUE %then %do;
        data tmp_excel_data_renamed;
            set tmp_excel_data_renamed;
            if cmiss(%do i = 1 %to &var_n;
                         &&var_new_name_&i
                         %if &i < &var_n %then %do; %bquote(,) %end;
                     %end;) = &var_n then delete;
        run;
    %end;

    /*清除所有变量的 format 和 informat*/
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

    /*创建输出数据集*/
    %if %bquote(&sort_by) ^= #NULL %then %do;
        proc sort data = tmp_excel_data_renamed out = &outdata;
            by &sort_by;
        run;
    %end;
    %else %do;
        data &outdata;
            set tmp_excel_data_renamed;
        run;
    %end;

    /*删除中间数据集*/
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
    %put NOTE: 宏程序 ads_import_excel 已结束运行！;
%mend;
