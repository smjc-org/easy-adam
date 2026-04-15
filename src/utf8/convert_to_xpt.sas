/*
 * Macro Name:    convert_to_xpt
 * Macro Purpose: sas7bdat -> xpt 批量转换
 * Author:        wtwang
 * Version Date:  2026-04-15
*/

%macro convert_to_xpt(sas7bdat_dir, xpt_dir, format = v8, debug = false) / parmbuff;
    /*  sas7bdat_dir: sas7bdat 文件夹
     *  xpt_dir:      xpt 文件夹
     *  format:       xpt 格式的版本，默认值 v8，可选 v5, v8, auto
     *  debug:        调试模式
    */

    /*统一参数大小写*/
    %let sas7bdat_dir = %sysfunc(strip(%bquote(&sas7bdat_dir)));
    %let xpt_dir      = %sysfunc(strip(%bquote(&xpt_dir)));
    %let format       = %upcase(%sysfunc(strip(%bquote(&format))));
    %let debug        = %upcase(%sysfunc(strip(%bquote(&debug))));

    /*确保目录存在*/
    filename fsas "&sas7bdat_dir";
    filename fxpt "&xpt_dir";

    %let fsas_id = %sysfunc(dopen(fsas));
    %let fxpt_id = %sysfunc(dopen(fxpt));

    %if &fsas_id = 0 %then %do;
        %put ERROR: %sysfunc(sysmsg());
        %goto exit;
    %end;

    %if &fxpt_id = 0 %then %do;
        X "mkdir ""&xpt_dir"" & exit";
    %end;

    /*建立临时逻辑库*/
    %let lib_sas7bdat = TMP_SAS;
    libname &lib_sas7bdat "&sas7bdat_dir";

    /*获取 sas7bdat 文件列表*/
    proc sql noprint;
        select memname into :sas7bdat_1- from dictionary.tables where libname = "&lib_sas7bdat" and memtype = "DATA";
        %let sas7bdat_n = &sqlobs;
    quit;

    /*批量调用 %loc2xpt*/
    %do i = 1 %to &sas7bdat_n;
        filename fsas_&i "&xpt_dir\&&sas7bdat_&i...xpt";
        %loc2xpt(filespec = fsas_&i, libref = &lib_sas7bdat, memlist = &&sas7bdat_&i, format = &format);
    %end;

    /*清除临时逻辑库*/
    libname &lib_sas7bdat clear;

    /*删除中间数据集*/
    %if %bquote(&debug) = %upcase(false) %then %do;

    %end;

    %exit:
    %let rc = %sysfunc(dclose(&fsas_id));
    %let rc = %sysfunc(dclose(&fxpt_id));
    filename fsas clear;
    filename fxpt clear;
    %put NOTE: 宏程序 convert_to_xpt 已结束运行！;
%mend;
