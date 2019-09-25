# utl-copy-all-sas-datasets-in-work-library-to-tabs-in-one-excel-workbook
Copy all sas datasets in work library to tabs in one excel workbook 

    Copy all sas datasets in work library to tabs in one excel workbook                                                           
                                                                                                                                  
    This does not use 'proc export';                                                                                              
                                                                                                                                  
    github                                                                                                                        
    https://tinyurl.com/y379fbcc                                                                                                  
    https://github.com/rogerjdeangelis/utl-copy-all-sas-datasets-in-work-library-to-tabs-in-one-excel-workbook                    
                                                                                                                                  
    SAS Forum                                                                                                                     
    https://tinyurl.com/y5nlxf5x                                                                                                  
    https://communities.sas.com/t5/SAS-Programming/How-to-export-all-dataset-present-in-work-library-to-an-excel/m-p/591482       
                                                                                                                                  
    *_                   _                                                                                                        
    (_)_ __  _ __  _   _| |_                                                                                                      
    | | '_ \| '_ \| | | | __|                                                                                                     
    | | | | | |_) | |_| | |_                                                                                                      
    |_|_| |_| .__/ \__,_|\__|                                                                                                     
            |_|                                                                                                                   
    ;                                                                                                                             
                                                                                                                                  
    * Make work datasets;                                                                                                         
                                                                                                                                  
    proc datasets lib=work kill;                                                                                                  
    run;quit;                                                                                                                     
                                                                                                                                  
    data one two tre;                                                                                                             
       set sashelp.class(keep=name sex age);                                                                                      
       select ;                                                                                                                   
            when (mod(_n_,3)=0) output one;                                                                                       
            when (mod(_n_,4)=0) output two;                                                                                       
            when (mod(_n_,5)=0) output tre;                                                                                       
            otherwise;                                                                                                            
       end;                                                                                                                       
    run;quit;                                                                                                                     
                                                                                                                                  
    Three datasets in work  library                                                                                               
                                                                                                                                  
          Member  Obs, Entries                                                                                                    
    Name  Type     or Indexes   Vars                                                                                              
                                                                                                                                  
    ONE   DATA         6         5                                                                                                
    TRE   DATA         2         5                                                                                                
    TWO   DATA         3         5                                                                                                
                                                                                                                                  
                                                                                                                                  
    WORK.ONE total obs=6                                                                                                          
                                                                                                                                  
     Name       Sex    Age                                                                                                        
                                                                                                                                  
     Barbara     F      13                                                                                                        
     James       M      12                                                                                                        
     Jeffrey     M      13                                                                                                        
     Judy        F      14                                                                                                        
     Philip      M      16                                                                                                        
     Thomas      M      11                                                                                                        
                                                                                                                                  
                                                                                                                                  
    WORK.TWO total obs=3                                                                                                          
                                                                                                                                  
      Name     Sex    Age                                                                                                         
                                                                                                                                  
     Carol      F      14                                                                                                         
     Janet      F      15                                                                                                         
     Robert     M      12                                                                                                         
                                                                                                                                  
                                                                                                                                  
    WORK.TRE total obs=2                                                                                                          
                                                                                                                                  
     Name     Sex    Age                                                                                                          
                                                                                                                                  
     Henry     M      14                                                                                                          
     John      M      12                                                                                                          
                                                                                                                                  
    *            _               _                                                                                                
      ___  _   _| |_ _ __  _   _| |_                                                                                              
     / _ \| | | | __| '_ \| | | | __|                                                                                             
    | (_) | |_| | |_| |_) | |_| | |_                                                                                              
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                             
                    |_|                                                                                                           
    ;                                                                                                                             
                                                                                                                                  
    ONE WORKBOOK WITH THREE TABS;                                                                                                 
    =============================                                                                                                 
                                                                                                                                  
    WORK.LOG total obs=3                                                                                                          
                                                                                                                                  
                  Condition                                                                                                       
      mem    rc    code      status                                                                                               
                                                                                                                                  
      ONE     0     0      ONE copied                                                                                             
      TRE     0     0      TRE copied                                                                                             
      TWO     0     0      TWO copied                                                                                             
                                                                                                                                  
                                                                                                                                  
    d:/xls/tabs.xlsx                                                                                                              
                                                                                                                                  
       +------------------+   +------------------+  +------------------+                                                          
       |  A    |  B  |  C |   |  A    |  B  |  C |  |  A    |  B  |  C |                                                          
       +------------------+   +------------------+  +------------------+                                                          
     1 |NAME   |AGE  |SEX |   |NAME   |AGE  |SEX |  |NAME   |AGE  |SEX |                                                          
       +-------+-----+----|   +-------+-----+----|  +-------+-----+----+                                                          
     2 |Alfred |14   |M   |   |Alice  |14   |F   |  |Barbara|14   |M   |                                                          
       +-------+-----+----+   +-------+-----+----+  --------------------                                                          
     3 |Carol  |13   |F   |   |Henry  |14   |M   |  |James  |13   |M   |                                                          
       -------------------+   --------------------  --------------------                                                          
       [ONE]                  [TWO]                 [TRE]                                                                         
                                                                                                                                  
                                                                                                                                  
    *                                                                                                                             
     _ __  _ __ ___   ___ ___  ___ ___                                                                                            
    | '_ \| '__/ _ \ / __/ _ \/ __/ __|                                                                                           
    | |_) | | | (_) | (_|  __/\__ \__ \                                                                                           
    | .__/|_|  \___/ \___\___||___/___/                                                                                           
    |_|                                                                                                                           
    ;                                                                                                                             
                                                                                                                                  
    * MAKE DATA;                                                                                                                  
                                                                                                                                  
    %utlfkil(d:/xls/tabs.xlsx);  * delete workbook if it exists;                                                                  
                                                                                                                                  
    proc datasets lib=work kill;                                                                                                  
    run;quit;                                                                                                                     
                                                                                                                                  
    data one two tre;                                                                                                             
       set sashelp.class(keep=name sex age);                                                                                      
       select ;                                                                                                                   
            when (mod(_n_,3)=0) output one;                                                                                       
            when (mod(_n_,4)=0) output two;                                                                                       
            when (mod(_n_,5)=0) output tre;                                                                                       
            otherwise;                                                                                                            
       end;                                                                                                                       
    run;quit;                                                                                                                     
                                                                                                                                  
                                                                                                                                  
    * SOLUTION;                                                                                                                   
                                                                                                                                  
    libname xel "d:/xls/tabs.xlsx";                                                                                               
                                                                                                                                  
    %symdel names cc nams /nowarn; * just in case;                                                                                
                                                                                                                                  
    data log ;                                                                                                                    
                                                                                                                                  
      if _n_=0 then do; %let rc=%sysfunc(                                                                                         
                                                                                                                                  
         dosubl('                                                                                                                 
             ods select members;                                                                                                  
             ods output members=dsns;                                                                                             
             proc contents data=work._all_ mt=data;                                                                               
             run;quit;                                                                                                            
             ods output close;                                                                                                    
                                                                                                                                  
             proc sql;                                                                                                            
              select quote(name) into :names separated by "," from dsns                                                           
             ;quit;                                                                                                               
             %let nams=&sqlobs;                                                                                                   
         '));                                                                                                                     
      end;                                                                                                                        
                                                                                                                                  
      do mem=&names;                                                                                                              
                                                                                                                                  
         call symputx('member',mem);                                                                                              
                                                                                                                                  
         rc=dosubl('                                                                                                              
            data xel.&member;                                                                                                     
                set &member;                                                                                                      
            run;quit;                                                                                                             
            %let cc=&syserr;                                                                                                      
         ');                                                                                                                      
                                                                                                                                  
         condition_code=symget("CC");                                                                                             
                                                                                                                                  
         if condition_code="0" then status=catx(" ",mem, "copied");                                                               
         else status=catx(" ",mem, "error");                                                                                      
                                                                                                                                  
         output;                                                                                                                  
                                                                                                                                  
      end;                                                                                                                        
                                                                                                                                  
    run;quit;                                                                                                                     
                                                                                                                                  
    libname xel clear;                                                                                                            
                                                                                                                                  
