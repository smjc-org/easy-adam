/*
 * Macro Name:    ads_rawdata_simulate_by_excel
 * Macro Purpose: 根据 Excel 数据库生成模拟原始数据
 * Author:        wtwang
 * Version Date:  2025-02-13 0.1.0
*/

%macro ads_rawdata_simulate_by_excel(file,
                                     outdata,
                                     sheet_name,
                                     range_attr,
                                     types,
                                     warning_var_name_empty = true,
                                     warning_var_name_not_meet_v7 = true,
                                     warning_var_name_len_gt_8 = true, debug = false) / parmbuff;
    /*  file:                         Excel 文件路径
     *  outdata:                      SAS 数据集名称
     *  sheet_name:                   工作表名称
     *  range_attr:                   包含变量名和标签定义的 2*C 单元格区域，其中第一行必须是标签，第二行必须是变量名
     *  types:                        变量类型的定义集合，例如：types = (usubjid = char(6), age = num(8))
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
    %let warning_var_name_empty       = %upcase(%sysfunc(strip(%bquote(&warning_var_name_empty))));
    %let warning_var_name_not_meet_v7 = %upcase(%sysfunc(strip(%bquote(&warning_var_name_not_meet_v7))));
    %let warning_var_name_len_gt_8    = %upcase(%sysfunc(strip(%bquote(&warning_var_name_len_gt_8))));
    %let debug                        = %upcase(%sysfunc(strip(%bquote(&debug))));

    /*读取 Excel 文件*/
    proc import file = "&file" out = tmp_excel_attr dbms = xlsx replace;
        sheet = "&sheet_name";
        range = "&range_attr";
        getnames = no;
    run;

    /*转置 tmp_excel_attr*/
    proc transpose data = tmp_excel_attr out = tmp_excel_attr_trans;
        var _all_;
    run;

    /*获取变量名称和变量标签*/
    proc sql noprint;
        select _NAME_       into :column_1-    from tmp_excel_attr_trans;
        select kstrip(COL2) into :var_name_1-  from tmp_excel_attr_trans;
        select kstrip(COL1) into :var_label_1- from tmp_excel_attr_trans;
    quit;
    %let var_n = &sqlobs;

    /*检查变量名是否符合规范*/
    %do i = 1 %to &var_n;
        %let IS_VALID_VAR_&i = TRUE;

        /*检查变量名是否为空*/
        %if &warning_var_name_empty = TRUE %then %do;
            %if %superq(var_name_&i) = %bquote() %then %do;
                %let IS_VALID_VAR_&i = FALSE;
                %put WARNING: 列 %superq(column_&i) 的变量名为空！;
            %end;
        %end;

        /*检查变量名是否在 VALIDVARNAME=V7 下合法*/
        %if &warning_var_name_not_meet_v7 = TRUE %then %do;
            %if %sysfunc(notname(%superq(var_name_&i))) > 0 %then %do;
                %let IS_VALID_VAR_&i = FALSE;
                %put WARNING: 列 %superq(column_&i) 的变量名 %superq(var_name_&i) 在 VALIDVARNAME=V7 下不合法！;
            %end;
        %end;

        /*检查变量名的长度是否超过 8*/
        %if &warning_var_name_len_gt_8 = TRUE %then %do;
            %if %length(%superq(var_name_&i)) > 8 %then %do;
                %put WARNING: 列 %superq(column_&i) 的变量名 %superq(var_name_&i) 的长度超过 8 ！;
            %end;
        %end;

        %if &&IS_VALID_VAR_&i = FALSE %then %do;
            %let var_name_&i = &&column_&i;
        %end;
    %end;

    /*创建模拟数据集*/
    proc sql noprint;
        create table tmp_rawdata_simulated
            (
                %do i = 1 %to &var_n;
                    &&var_name_&i char(200) label = %unquote(%str(%')%superq(var_label_&i)%str(%'))
                    %if &i < &var_n %then %do; %bquote(,) %end;
                %end;
            );
    quit;

    /*创建输出数据集*/
    data &outdata;
        set tmp_rawdata_simulated;
    run;

    /*删除中间数据集*/
    %if %bquote(&debug) = %upcase(false) %then %do;
        proc datasets library = work nowarn noprint;
            delete tmp_excel_attr
                   tmp_excel_attr_trans
                   tmp_rawdata_simulated
                   ;
        quit;
    %end;

    %exit:
    %put NOTE: 宏程序 ads_import_excel 已结束运行！;
%mend;


%ads_rawdata_simulate_by_excel(file = %str(D:\OneDrive\统计部\项目\MD\2023\04 赛诺威盛-OmegeCT One X射线计算机体层摄影设备\03 数据管理\06 数据库\V0.3\数据库_赛诺威盛_OmegaCT One X 射线计算机体层摄影设备_V0.3_20250211.xlsx),
                               outdata = info,
                               sheet_name = %str(基本信息(INFO)),
                               range_attr = %str(A1:U2));


