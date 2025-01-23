/*
 * Macro Name:    ads_empty
 * Macro Purpose: ��ȡ ADaM Specification�����������ݼ�
 * Author:        wtwang
 * Version Date:  2025-01-23 0.1.0
*/

%macro ads_empty(spec, prefix = %str(), suffix = %str(_empty), sheet_name = %str(����˵��), range = #ALL, debug = false) / parmbuff;
    /*  spec:       ADaM Specification �ļ�·��
     *  prefix:     �����ݼ�������ǰ׺
     *  suffix:     �����ݼ������ƺ�׺
     *  sheet_name: ������������Ĺ���������
     *  range:      sheet_name �������е��ض���Χ�����磺A1$F255
     *  debug:      ����ģʽ
    */

    /*ͳһ������Сд*/
    %let spec       = %sysfunc(strip(%bquote(&spec)));
    %let prefix     = %upcase(%sysfunc(strip(%bquote(&prefix))));
    %let suffix     = %upcase(%sysfunc(strip(%bquote(&suffix))));
    %let sheet_name = %sysfunc(strip(%bquote(&sheet_name)));
    %let range      = %upcase(%sysfunc(strip(%bquote(&range))));
    %let debug      = %upcase(%sysfunc(strip(%bquote(&debug))));

    /*��ȡ ADaM Specification*/
    proc import file = "&spec" out = tmp_adam_spec dbms = xlsx replace;
        sheet = "&sheet_name";
        %if %bquote(&range) ^= %bquote(#ALL) %then %do;
            range = "&range";
        %end;
    run;

    /*����������ǿ�ɶ���*/
    proc datasets library = work nowarn noprint;
        modify tmp_adam_spec;
            rename VAR1 = dataset
                   VAR2 = variable
                   VAR3 = label
                   VAR4 = type
                   VAR5 = length
                   VAR6 = format
                   VAR7 = terminology
                   VAR8 = core
                   VAR9 = source
                   VAR10 = details;
    quit;

    /*��ȡ���ݼ�����*/
    proc sql noprint;
        select distinct dataset into :dataset_name_1- from tmp_adam_spec where not missing(dataset);
    quit;
    %let dataset_n = &sqlobs;

    /*���������ݼ�*/
    %do i = 1 %to &dataset_n;
        proc sql noprint;
            create table tmp_adam_spec_&&dataset_name_&i as select * from tmp_adam_spec where dataset = "&&dataset_name_&i" and not missing(variable);

            /*����Ƿ����ظ�����ı���*/
            select count(*) into :n from tmp_adam_spec_&&dataset_name_&i;
            select count(distinct variable) into :varn from tmp_adam_spec_&&dataset_name_&i;

            %if &varn < &n %then %do;
                select variable into :dup_var_1- from tmp_adam_spec_&&dataset_name_&i group by variable having count(*) >= 2;
                %let dup_var_list_expr = %bquote();
                %do j = 1 %to &sqlobs;
                    %if &j = 1 %then %let dup_var_list_expr = %bquote(&&dup_var_&j);
                    %else %let dup_var_list_expr = &dup_var_list_expr%bquote(,) %bquote(&&dup_var_&j);
                %end;
                %put ERROR: ���ݼ� &&dataset_name_&i �д����ظ�����ı��� &dup_var_list_expr ��;
                %goto exit;
            %end;

            /*���� SQL ���*/
            select variable into :var_1-    from tmp_adam_spec_&&dataset_name_&i;
            select label    into :label_1-  from tmp_adam_spec_&&dataset_name_&i;
            select type     into :type_1-   from tmp_adam_spec_&&dataset_name_&i;
            select length   into :length_1- from tmp_adam_spec_&&dataset_name_&i;
            select format   into :format_1- from tmp_adam_spec_&&dataset_name_&i;
            create table &prefix.&&dataset_name_&i..&suffix
                (
                    %do j = 1 %to &varn;
                        %if %bquote(&&type_&j) = %bquote() %then %do;
                            %put WARNING: ���ݼ� &&dataset_name_&i �б��� %bquote(&&var_&j) ������Ϊ�գ�������Ϊ�ַ��ͣ�;
                            %let type_&j = %bquote(CHAR);
                        %end;
                        %else %if %upcase(&&type_&j) ^= %bquote(CHAR) and %upcase(&&type_&j) ^= %bquote(NUM) %then %do;
                            %put WARNING: ���ݼ� &&dataset_name_&i �б��� %bquote(&&var_&j) ������ %bquote(&&type_&j) ���Ϸ���������Ϊ�ַ��ͣ�;
                            %let type_&j = %bquote(CHAR);
                        %end;

                        %if %bquote(&&label_&j) = %bquote() %then %do;
                            %let label_&j = %bquote(&&var_&j);
                            %put WARNING: ���ݼ� &&dataset_name_&i �б��� %bquote(&&var_&j) �ı�ǩΪ�գ����Զ�����Ϊ��������;
                        %end;

                        &&var_&j
                        %if %upcase(&&type_&j) = %bquote(NUM) %then %do;
                            %if %bquote(&&length_&j) = %bquote() %then %do;
                                %let length_&j = 8;
                                %put WARNING: ���ݼ� &&dataset_name_&i �б��� %bquote(&&var_&j) �ĳ���Ϊ�գ����Զ�����Ϊ 8 ��;
                            %end;
                            numeric(&&length_&j)
                        %end;
                        %else %if %upcase(&&type_&j) = %bquote(CHAR) %then %do;
                            %if %bquote(&&length_&j) = %bquote() %then %do;
                                %let length_&j = 200;
                                %put WARNING: ���ݼ� &&dataset_name_&i �б��� %bquote(&&var_&j) �ĳ���Ϊ�գ����Զ�����Ϊ 200 ��;
                            %end;
                            char(&&length_&j)
                        %end;
                        %else %do;
                            %put ERROR: ������Ԥ��֮��Ĵ���;
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

    /*ɾ���м����ݼ�*/
    %if %bquote(&debug) = %upcase(false) %then %do;
        proc datasets library = work nowarn noprint;
            delete tmp_adam_spec
                   %do i = 1 %to &dataset_n;
                       tmp_adam_spec_&&dataset_name_&i
                   %end;
                   ;
        quit;
    %end;

    %exit:
    %put NOTE: ����� ads_empty �ѽ������У�;
%mend;
