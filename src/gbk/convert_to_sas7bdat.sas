/*
 * Macro Name:    convert_to_sas7bdat
 * Macro Purpose: xpt -> sas7bdat 批量转换
 * Author:        wtwang
 * Version Date:  2026-04-15
*/

%macro convert_to_sas7bdat(dir, libref = work) / parmbuff;
    /*  dir:          xpt 文件所在目录的路径
     *  libref:       存放 sas7bdat 文件的逻辑库名称，默认值 work
    */

    /*统一参数大小写*/
    %let dir     = %sysfunc(strip(%bquote(&dir)));
    %let libref  = %sysfunc(strip(%bquote(&libref)));

    /*确保目录存在*/
    filename fxpt "&dir";

    %let fxpt_id = %sysfunc(dopen(fxpt));
    
    %if &fxpt_id = 0 %then %do;
        %put ERROR: %sysfunc(sysmsg());
        %goto exit;
    %end;

    /*获取 xpt 文件列表*/
    filename dinfo pipe "cd /d ""&dir"" & dir *.xpt /b";
    data _null_;
        infile dinfo truncover end = end;
        input xpt $200.;
        name = scan(xpt, 1, ".");
        call symputx(cats('xpt_', _n_), name);

        if end then call symputx('xpt_n', _n_);
    run;

    /*批量调用 %xpt2loc*/
    %do i = 1 %to &xpt_n;
        %put NOTE: 正在转换 &&xpt_&i...xpt -> &&xpt_&i...sas7bdat;
        options nonotes;
        filename fxpt_&i "&dir\&&xpt_&i...xpt";
        %xpt2loc(filespec = fxpt_&i, libref = %bquote(&libref));
        options notes;
    %end;

    %exit:
    %let rc = %sysfunc(dclose(&fxpt_id));
    filename fxpt clear;
    %put NOTE: 宏程序 convert_to_sas7bdat 已结束运行！;
%mend;



