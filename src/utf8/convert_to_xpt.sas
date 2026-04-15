/*
 * Macro Name:    convert_to_xpt
 * Macro Purpose: sas7bdat -> xpt 批量转换
 * Author:        wtwang
 * Version Date:  2026-04-15
*/

%macro convert_to_xpt(libref, dir, format = v8, debug = false) / parmbuff;
    /*  libref: 存放 sas7bdat 文件的逻辑库名称
     *  dir:    xpt 文件所在目录的路径
     *  format: xpt 格式的版本，默认值 v8，可选 v5, v8, auto
    */

    /*统一参数大小写*/
    %let libref = %upcase(%sysfunc(strip(%bquote(&libref))));
    %let dir    = %sysfunc(strip(%bquote(&dir)));
    %let format = %upcase(%sysfunc(strip(%bquote(&format))));

    /*确保目录存在*/
    filename fxpt "&dir";

    %let fxpt_id = %sysfunc(dopen(fxpt));

    %if &fxpt_id = 0 %then %do;
        X "mkdir ""&dir"" & exit";
    %end;

    /*获取 sas7bdat 文件列表*/
    proc sql noprint;
        select memname into :sas7bdat_1- from dictionary.tables where libname = "&libref" and memtype = "DATA";
        %let sas7bdat_n = &sqlobs;
    quit;

    /*批量调用 %loc2xpt*/
    %do i = 1 %to &sas7bdat_n;
        %put NOTE: 正在转换 &&sas7bdat_&i...sas7bdat -> &&sas7bdat_&i...xpt;
        options nonotes;
        filename fxpt_&i "&dir\&&sas7bdat_&i...xpt";
        %loc2xpt(filespec = fxpt_&i, libref = %bquote(&libref), memlist = &&sas7bdat_&i, format = &format);
        options notes;
    %end;

    %exit:
    %let rc = %sysfunc(dclose(&fxpt_id));
    filename fxpt clear;
    %put NOTE: 宏程序 convert_to_xpt 已结束运行！;
%mend;
