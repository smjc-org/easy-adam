/*
 * Macro Name:    ads_empty
 * Macro Purpose: 读取 ADaM Specification，创建空数据集
 * Author:        wtwang
 * Version Date:  2025-01-23 0.1.0
                  2025-01-24 0.2.0
*/

%macro ads_empty(spec, select = #ALL, prefix = %str(), suffix = %str(_empty), sheet_name = %str(变量说明), range = #ALL, debug = false) / parmbuff;
    /*  spec:       ADaM Specification 文件路径
     *  select:     选择要输出的 ADaM 数据集，#ALL 表示所有数据集
     *  prefix:     空数据集的名称前缀
     *  suffix:     空数据集的名称后缀
     *  sheet_name: 包含变量定义的工作表名称
     *  range:      sheet_name 工作表中的特定范围，例如：A1$F255
     *  debug:      调试模式
    */

    /*统一参数大小写*/
    %let spec       = %sysfunc(strip(%bquote(&spec)));
    %let select     = %upcase(%sysfunc(strip(%bquote(&select))));
    %let prefix     = %upcase(%sysfunc(strip(%bquote(&prefix))));
    %let suffix     = %upcase(%sysfunc(strip(%bquote(&suffix))));
    %let sheet_name = %sysfunc(strip(%bquote(&sheet_name)));
    %let range      = %upcase(%sysfunc(strip(%bquote(&range))));
    %let debug      = %upcase(%sysfunc(strip(%bquote(&debug))));

    /*读取 ADaM Specification*/
    proc import file = "&spec" out = tmp_adam_spec dbms = xlsx replace;
        sheet = "&sheet_name";
        %if %bquote(&range) ^= %bquote(#ALL) %then %do;
            range = "&range";
        %end;
    run;

    /*重命名，增强可读性*/
    proc datasets library = work nowarn noprint;
        modify tmp_adam_spec;
            rename VAR1 = dataset
                   VAR2 = variable
                   VAR3 = label
                   VAR4 = type
                   VAR5 = length
                   VAR6 = format;
    quit;

    /*获取数据集名称*/
    proc sql noprint;
        select distinct upcase(dataset) into :dataset_1- from tmp_adam_spec where not missing(dataset);
    quit;
    %let dataset_n = &sqlobs;

    /*获取被选定的数据集名称*/
    %if %bquote(&select) = %bquote(#ALL) %then %do;
        %let dataset_selected_n = &dataset_n;
        %do i = 1 %to &dataset_selected_n;
            %let dataset_selected_&i = &&dataset_&i;
        %end;
    %end;
    %else %do;
        %let dataset_selected_n = %sysfunc(countw(%bquote(&select), %bquote(), %bquote(s)));
        %do i = 1 %to &dataset_selected_n;
            %let dataset_selected_&i = %scan(%bquote(&select), &i, %bquote(), %bquote(s));
            /*检查选定的数据集是否存在定义*/
            %let dataset_selected_&i._defined = false;
            %do j = 1 %to &dataset_n;
                %if &&dataset_selected_&i = &&dataset_&j %then %do;
                    %let dataset_selected_&i._defined = true;
                %end;
            %end;
            %if &&dataset_selected_&i._defined = false %then %do;
                %put WARNING: 找不到数据集 &&dataset_selected_&i 的变量定义！;
            %end;
        %end;
    %end;

    /*创建空数据集*/
    %do i = 1 %to &dataset_selected_n;
        %if &&dataset_selected_&i._defined = true %then %do;
            proc sql noprint;
                create table tmp_adam_spec_&&dataset_selected_&i as select * from tmp_adam_spec where upcase(dataset) = "&&dataset_selected_&i" and not missing(variable);

                /*检查是否有重复定义的变量*/
                select count(*) into :n from tmp_adam_spec_&&dataset_selected_&i;
                select count(distinct variable) into :varn from tmp_adam_spec_&&dataset_selected_&i;

                %if &varn < &n %then %do;
                    select variable into :dup_var_1- from tmp_adam_spec_&&dataset_selected_&i group by variable having count(*) >= 2;
                    %let dup_var_list_expr = %bquote();
                    %do j = 1 %to &sqlobs;
                        %if &j = 1 %then %let dup_var_list_expr = %bquote(&&dup_var_&j);
                        %else %let dup_var_list_expr = &dup_var_list_expr%bquote(,) %bquote(&&dup_var_&j);
                    %end;
                    %put ERROR: 数据集 &&dataset_selected_&i 中存在重复定义的变量 &dup_var_list_expr ！;
                    %goto exit;
                %end;

                /*构建 SQL 语句*/
                select variable into :var_1-    from tmp_adam_spec_&&dataset_selected_&i;
                select label    into :label_1-  from tmp_adam_spec_&&dataset_selected_&i;
                select type     into :type_1-   from tmp_adam_spec_&&dataset_selected_&i;
                select length   into :length_1- from tmp_adam_spec_&&dataset_selected_&i;
                select format   into :format_1- from tmp_adam_spec_&&dataset_selected_&i;
                create table &prefix.&&dataset_selected_&i..&suffix
                    (
                        %do j = 1 %to &varn;
                            %if %bquote(&&type_&j) = %bquote() %then %do;
                                %put WARNING: 数据集 &&dataset_selected_&i 中变量 %bquote(&&var_&j) 的类型为空，已重置为字符型！;
                                %let type_&j = %bquote(CHAR);
                            %end;
                            %else %if %upcase(&&type_&j) ^= %bquote(CHAR) and %upcase(&&type_&j) ^= %bquote(NUM) %then %do;
                                %put WARNING: 数据集 &&dataset_selected_&i 中变量 %bquote(&&var_&j) 的类型 %bquote(&&type_&j) 不合法，已重置为字符型！;
                                %let type_&j = %bquote(CHAR);
                            %end;

                            %if %bquote(&&label_&j) = %bquote() %then %do;
                                %let label_&j = %bquote(&&var_&j);
                                %put WARNING: 数据集 &&dataset_selected_&i 中变量 %bquote(&&var_&j) 的标签为空，已自动设置为变量名！;
                            %end;

                            &&var_&j
                            %if %upcase(&&type_&j) = %bquote(NUM) %then %do;
                                %if %bquote(&&length_&j) = %bquote() %then %do;
                                    %let length_&j = 8;
                                    %put WARNING: 数据集 &&dataset_selected_&i 中变量 %bquote(&&var_&j) 的长度为空，已自动设置为 8 ！;
                                %end;
                                numeric(&&length_&j)
                            %end;
                            %else %if %upcase(&&type_&j) = %bquote(CHAR) %then %do;
                                %if %bquote(&&length_&j) = %bquote() %then %do;
                                    %let length_&j = 200;
                                    %put WARNING: 数据集 &&dataset_selected_&i 中变量 %bquote(&&var_&j) 的长度为空，已自动设置为 200 ！;
                                %end;
                                char(&&length_&j)
                            %end;
                            %else %do;
                                %put ERROR: 发生了预料之外的错误！;
                                %goto exit;
                            %end;
                            label  = %unquote(%str(%')%bquote(&&label_&j)%str(%'))

                            %if %bquote(&&format_&j) ^= %bquote() %then %do;
                                format = %bquote(&&format_&j)
                            %end;

                            %if &j < &varn %then %do;
                                %bquote(,)
                            %end;
                        %end;
                    );
            quit;
        %end;
    %end;

    /*删除中间数据集*/
    %if %bquote(&debug) = %upcase(false) %then %do;
        proc datasets library = work nowarn noprint;
            delete tmp_adam_spec
                   %do i = 1 %to &dataset_selected_n;
                       tmp_adam_spec_&&dataset_selected_&i
                   %end;
                   ;
        quit;
    %end;

    %exit:
    %put NOTE: 宏程序 ads_empty 已结束运行！;
%mend;
