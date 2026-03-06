/*
 * Macro Name:    ads_compare_dir
 * Macro Purpose: 比较两个文件夹下的 analysis datasets
 * Author:        wtwang
 * Version Date:  2026-03-06
*/

%macro ads_compare_dir(base_dir,
                       compare_dir,
                       outdata,
                       debug = false) / parmbuff;
    /*  base_dir:    基准文件夹
     *  compare_dir: 比较文件夹
     *  outdata:     存储比较结果的数据集名称
     *  debug:       调试模式
    */

    /*统一参数大小写*/
    %let base_dir    = %sysfunc(strip(%bquote(&base_dir)));
    %let compare_dir = %sysfunc(strip(%bquote(&compare_dir)));
    %let outdata     = %sysfunc(strip(%bquote(&outdata)));
    %let debug       = %upcase(%sysfunc(strip(%bquote(&debug))));

    /*建立临时逻辑库*/
    %let libname_base    = TMP_B;
    %let libname_compare = TMP_C;
    libname &libname_base    "&base_dir";
    libname &libname_compare "&compare_dir";

    /*构建成员名对照表*/
    proc sql noprint;
        create table tmp_member_base    as select * from dictionary.members where libname = "&libname_base";
        create table tmp_member_compare as select * from dictionary.members where libname = "&libname_compare";
        create table tmp_map as
            select
                b.libname as b_libname,
                b.memname as b_memname,
                c.libname as c_libname,
                c.memname as c_memname
            from tmp_member_base as b full join tmp_member_compare as c on b.memname = c.memname
            order by b_memname, c_memname;
    quit;

    /*筛选两个逻辑库中都存在的数据集进行比较*/
    proc sql noprint;
        select b_memname into :memname_list separated by "," from tmp_map where b_memname = c_memname;
        %let memname_list_n = &sqlobs;
    quit;

    options nonotes;
    %do i = 1 %to &memname_list_n;
        %let memname_&i = %scan(%bquote(&memname_list), &i, %bquote(,));
        proc sql noprint;
            select nlobs into :b_nlobs trimmed from dictionary.tables where libname = "&libname_base"    and memname = "&&memname_&i";
            select nlobs into :c_nlobs trimmed from dictionary.tables where libname = "&libname_compare" and memname = "&&memname_&i";
        quit;

        %if &b_nlobs = 0 and &c_nlobs = 0 %then %do;
            %let &&memname_&i.._sysinfo = -1;
        %end;
        %else %if &b_nlobs = 0 and &c_nlobs > 0 %then %do;
            %let &&memname_&i.._sysinfo = -2;
        %end;
        %else %if &b_nlobs > 0 and &c_nlobs = 0 %then %do;
            %let &&memname_&i.._sysinfo = -3;
        %end;
        %else %do;
            proc compare base = &libname_base..&&memname_&i compare = &libname_compare..&&memname_&i noprint;
            run;
            %let &&memname_&i.._sysinfo = &sysinfo;
        %end;
    %end;
    options notes;

    /*整理比较结果*/
    options nonotes;
    proc sql noprint;
        create table tmp_result as select * from tmp_map;
        alter table tmp_result
            add sysinfo     num(8),
                sysinfo_p0 char(50),
                sysinfo_p1 char(50),
                sysinfo_p2 char(50),
                sysinfo_p3 char(50),
                sysinfo_p4 char(50),
                sysinfo_p5 char(50),
                sysinfo_p6 char(50),
                sysinfo_p7 char(50),
                sysinfo_p8 char(50),
                sysinfo_p9 char(50),
                sysinfo_p10 char(50),
                sysinfo_p11 char(50),
                sysinfo_p12 char(50),
                sysinfo_p13 char(50),
                sysinfo_p14 char(50),
                sysinfo_p15 char(50),
                result      char(1000),
                comment     char(200);
        %do i = 1 %to &memname_list_n;
            update tmp_result
                set sysinfo = %superq(&&memname_&i.._sysinfo) where b_memname = "&&memname_&i";
        %end;
        update tmp_result
            set sysinfo_p0  = ifc(band(sysinfo, 2**0),  "数据集标签不一致", ""),
                sysinfo_p1  = ifc(band(sysinfo, 2**1),  "数据集类型不一致", ""),
                sysinfo_p2  = ifc(band(sysinfo, 2**2),  "存在输入格式不同的变量", ""),
                sysinfo_p3  = ifc(band(sysinfo, 2**3),  "存在输出格式不同的变量", ""),
                sysinfo_p4  = ifc(band(sysinfo, 2**4),  "存在长度不同的变量", ""),
                sysinfo_p5  = ifc(band(sysinfo, 2**5),  "存在标签不同的变量", ""),
                sysinfo_p6  = ifc(band(sysinfo, 2**6),  "base 数据集存在 compare 数据集没有的观测", ""),
                sysinfo_p7  = ifc(band(sysinfo, 2**7),  "compare 数据集存在 base 数据集没有的观测", ""),
                sysinfo_p8  = ifc(band(sysinfo, 2**8),  "base 数据集存在 compare 数据集没有的 BY 组", ""),
                sysinfo_p9  = ifc(band(sysinfo, 2**9),  "compare 数据集存在 base 数据集没有的 BY 组", ""),
                sysinfo_p10 = ifc(band(sysinfo, 2**10), "base 数据集存在 compare 数据集没有的变量", ""),
                sysinfo_p11 = ifc(band(sysinfo, 2**11), "compare 数据集存在 base 数据集没有的变量", ""),
                sysinfo_p12 = ifc(band(sysinfo, 2**12), "存在不同值", ""),
                sysinfo_p13 = ifc(band(sysinfo, 2**13), "变量类型冲突", ""),
                sysinfo_p14 = ifc(band(sysinfo, 2**14), "BY 变量不匹配", ""),
                sysinfo_p15 = ifc(band(sysinfo, 2**15), "致命错误：未进行比较", "");
    quit;
    data tmp_result;
        set tmp_result;
        select;
            when (not missing(b_memname) and missing(c_memname)) result  = "compare 中不存在";
            when (missing(b_memname) and not missing(c_memname)) result  = "base 中不存在";
            when (sysinfo = -1)                                  comment = "base 和 compare 数据集均为空";
            when (sysinfo = -2)                                  comment = "base 数据集为空";
            when (sysinfo = -3)                                  comment = "compare 数据集为空";
            otherwise result = catq("DT", ", ", sysinfo_p0,
                                                sysinfo_p1,
                                                sysinfo_p2,
                                                sysinfo_p3,
                                                sysinfo_p4,
                                                sysinfo_p5,
                                                sysinfo_p6,
                                                sysinfo_p7,
                                                sysinfo_p8,
                                                sysinfo_p9,
                                                sysinfo_p10,
                                                sysinfo_p11,
                                                sysinfo_p12,
                                                sysinfo_p13,
                                                sysinfo_p14,
                                                sysinfo_p15);
        end;
    run;
    options notes;

    /*输出数据集*/
    proc sql noprint;
        create table &outdata as
            select
                b_memname label = "基准数据集",
                c_memname label = "比较数据集",
                result    label = "结果",
                comment   label = "备注"
            from tmp_result;
    quit;

    /*删除中间数据集*/
    %if %bquote(&debug) = %upcase(false) %then %do;
        proc datasets library = work nowarn noprint;
            delete tmp_member_base
                   tmp_member_compare
                   tmp_map
                   tmp_result
                   ;
        quit;
    %end;

    %exit:
    %put NOTE: 宏程序 ads_import_excel 已结束运行！;
%mend;
