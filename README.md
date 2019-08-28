# utl-ods-export-sas-table-to-execel-with-rotated-column-headers
ODS Export sas table to execel with rotated column headers
    ODS Export sas table to execel with rotated column headers                                                                          
                                                                                                                                        
    github                                                                                                                                                                                                                                  
    https://github.com/rogerjdeangelis/utl-ods-export-sas-table-to-excel-with-rotated-column-headers                       
                                                                                                                                        
    SAS forums                                                                                                                          
    https://tinyurl.com/y5mnnv2r                                                                                                        
    https://communities.sas.com/t5/SAS-Programming/Nested-column-names-different-sub-column-names-per-head-column/m-p/584531            
                                                                                                                                        
    https://tinyurl.com/y6sjxjch                                                                                                        
    https://communities.sas.com/t5/ODS-and-Base-Reporting/How-to-print-columns-vertically-using-SAS-Proc-Report-and-ODS/td-p/241905     
                                                                                                                                        
    ODS template to rotate excel header                                                                                                 
    http://support.sas.com/rnd/base/ods/odsmarkup/msoffice2k/index.html                                                                 
                                                                                                                                        
    KSharpe profile                                                                                                                     
    https://communities.sas.com/t5/user/viewprofilepage/user-id/18408                                                                   
                                                                                                                                        
    Somewhate dated paper                                                                                                               
    http://support.sas.com/resources/papers/proceedings09/223-2009.pdf                                                                  
                                                                                                                                        
    * as a side note SAS graph supports rotated text (anywhere?);                                                                       
                                                                                                                                        
    *_                   _                                                                                                              
    (_)_ __  _ __  _   _| |_                                                                                                            
    | | '_ \| '_ \| | | | __|                                                                                                           
    | | | | | |_) | |_| | |_                                                                                                            
    |_|_| |_| .__/ \__,_|\__|                                                                                                           
            |_|                                                                                                                         
    ;                                                                                                                                   
                                                                                                                                        
    Download template code 'msoffice2k_x.sas' from:                                                                                     
                                                                                                                                        
    https://communities.sas.com/t5/user/viewprofilepage/user-id/18408                                                                   
                                                                                                                                        
    data have;                                                                                                                          
      set sashelp.class;                                                                                                                
    run;quit;                                                                                                                           
                                                                                                                                        
                                                                                                                                        
     WORK.HAVE total obs=19                                                                                                             
                                                                                                                                        
      NAME       SEX    AGE    HEIGHT    WEIGHT                                                                                         
                                                                                                                                        
      Alfred      M      14     69.0      112.5                                                                                         
      Alice       F      13     56.5       84.0                                                                                         
      Barbara     F      13     65.3       98.0                                                                                         
    ...                                                                                                                                 
                                                                                                                                        
    *            _               _                                                                                                      
      ___  _   _| |_ _ __  _   _| |_                                                                                                    
     / _ \| | | | __| '_ \| | | | __|                                                                                                   
    | (_) | |_| | |_| |_) | |_| | |_                                                                                                    
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                                   
                    |_|                                                                                                                 
    ;                                                                                                                                   
                                                                                                                                        
      d:/xls/classx.xlsx                                                                                                                
                                                                                                                                        
       +----------------------------------------+                                                                                       
       |  A    |  B    |  C    |  D    |  E     |                                                                                       
       +----------------------------------------+                                                                                       
       |       |       |       |       |        |                                                                                       
       |       |       |       |  H    |  W     |                                                                                       
       |       |       |       |  E    |  E     |                                                                                       
     1 |   N   |       |       |  I    |  I     |                                                                                       
       |   A   |  A    |  S    |  G    |  G     |                                                                                       
       |   M   |  G    |  E    |  H    |  H     |                                                                                       
       |   E   |  E    |  X    |  T    |  T     |                                                                                       
       |       |       |       |       |        |                                                                                       
       |-------+-------+-------+-------+--------|                                                                                       
     2 |Alfred |14     |M      |69     |112.5   |                                                                                       
       |-------+-------+-------+-------+--------+                                                                                       
     3 |Alice  |13     |F      |56.5   |84      |                                                                                       
       |-------+-------+-------+-------+--------+                                                                                       
     4 |Barbara|13     |F      |65.3   |98      |                                                                                       
       |-------+-------+-------+-------+--------+                                                                                       
     5 |Carol  |14     |F      |62.8   |102.5   |                                                                                       
       |-------+-------+-------+-------+--------+                                                                                       
     6 |Henry  |14     |M      |63.5   |102.5   |                                                                                       
       ------------------------------------------                                                                                       
       ...                                                                                                                              
       [class]                                                                                                                          
                                                                                                                                        
    *                                                                                                                                   
     _ __  _ __ ___   ___ ___  ___ ___                                                                                                  
    | '_ \| '__/ _ \ / __/ _ \/ __/ __|                                                                                                 
    | |_) | | | (_) | (_|  __/\__ \__ \                                                                                                 
    | .__/|_|  \___/ \___\___||___/___/                                                                                                 
    |_|                                                                                                                                 
    ;                                                                                                                                   
                                                                                                                                        
    title;footnote;                                                                                                                     
    %utlfkil(d:\xls\classx.xls); * just in case it exists;                                                                              
                                                                                                                                        
    %include "c:/oto/msoffice2k_x.sas";                                                                                                 
                                                                                                                                        
    ods tagsets.msoffice2k_x file="d:\xls\classx.xls" style=normal                                                                      
        options( rotate_headers="90" height="60");                                                                                      
                                                                                                                                        
    proc print data=sashelp.class(obs=3);                                                                                               
    run;                                                                                                                                
                                                                                                                                        
    ods tagsets.msoffice2k_x close;                                                                                                     
                                                                                                                                        
    * right click and open with excel;                                                                                                  
    * variant of xml so the extension will issue a warning if clicked on directly;                                                      
                                                                                                                                        
                                                                                                                                        
                                                                                                                                        
