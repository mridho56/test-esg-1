*** ----------------------------------------------------------------------- ***
***              Load the dataset and create a STATA dataset                ***
*** ----------------------------------------------------------------------- ***

* Here goes your directory
cd "X:\My Documents\Test esg"
 
* clear the dataset previously loaded 
clear
set maxvar 15000

*** ----------------------Import the Enviromental Score--------------------------- ***
** Import the first part
import excel using Test_esg_2.xlsx, sheet("envscore") cellrange(B1:GJJ23) firstrow clear
 
gen Years = _n  

// Identify and convert string variables to numeric
ds SC*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist SC* {
    recast double `var'
}

// Verify the changes
describe SC*

 
* reshape from wide to long 
reshape long SC, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_SC_1, replace

** Import the second part
clear
import excel using Test_esg_2.xlsx, sheet("envscore") cellrange(GJK1:NTR23) firstrow clear

gen Years = _n  

// Identify and convert string variables to numeric
ds SC*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist SC* {
    recast double `var'
}

// Verify the changes
describe SC*

* reshape from wide to long 
reshape long SC, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_SC_2, replace

** Import the third part
clear
import excel using Test_esg_2.xlsx, sheet("envscore") cellrange(NTS1:UKN23) firstrow clear
 
gen Years = _n  
 
// Identify and convert string variables to numeric
ds SC*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist SC* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long SC, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_SC_3, replace

** Append the three parts
clear
append using Data_SC_1 Data_SC_2 Data_SC_3
save Data_SC, replace

* Save the dataset
save Data_SC, replace

*** ----------------------Import the Scope1 --------------------------- ***
** Import the first part
clear
import excel using Test_esg_2.xlsx, sheet("scope1") cellrange(B1:GJJ23) firstrow clear
 
gen Years = _n  

// Identify and convert string variables to numeric
ds SCA*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist SCA* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long SCA, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_SCA_1, replace

** Import the second part
clear
import excel using Test_esg_2.xlsx, sheet("scope1") cellrange(GJK1:NTR23) firstrow clear
 
gen Years = _n 
 
// Identify and convert string variables to numeric
ds SCA*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist SCA* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long SCA, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_SCA_2, replace

** Import the third part
clear
import excel using Test_esg_2.xlsx, sheet("scope1") cellrange(NTS1:UKN23) firstrow clear

gen Years = _n 
 
// Identify and convert string variables to numeric
ds SCA*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist SCA* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long SCA, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_SCA_3, replace

** Append the three parts
clear
append using Data_SCA_1 Data_SCA_2 Data_SCA_3
save Data_SCA, replace

* Save the dataset
save Data_SCA, replace

*** ----------------------Import the Scope2 --------------------------- ***
** Import the first part
clear
import excel using Test_esg_2.xlsx, sheet("scope2") cellrange(B1:GJJ23) firstrow clear
 
gen Years = _n  

// Identify and convert string variables to numeric
ds SCB*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist SCB* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long SCB, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_SCB_1, replace

** Import the second part
clear
import excel using Test_esg_2.xlsx, sheet("scope2") cellrange(GJK1:NTR22) firstrow clear
 
gen Years = _n 
 
// Identify and convert string variables to numeric
ds SCB*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist SCB* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long SCB, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_SCB_2, replace

** Import the third part
clear
import excel using Test_esg_2.xlsx, sheet("scope2") cellrange(NTS1:UKN23) firstrow clear

gen Years = _n 
 
// Identify and convert string variables to numeric
ds SCB*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist SCB* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long SCB, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_SCB_3, replace

** Append the three parts
clear
append using Data_SCB_1 Data_SCB_2 Data_SCB_3
save Data_SCB, replace

* Save the dataset
save Data_SCB, replace

*** ----------------------Import the Scope3 --------------------------- ***
** Import the first part
clear
import excel using Test_esg_2.xlsx, sheet("scope3") cellrange(B1:GJJ23) firstrow clear
 
gen Years = _n  

// Identify and convert string variables to numeric
ds SCC*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist SCC* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long SCC, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_SCC_1, replace

** Import the second part
clear
import excel using Test_esg_2.xlsx, sheet("scope3") cellrange(GJK1:NTR23) firstrow clear
 
gen Years = _n 
 
// Identify and convert string variables to numeric
ds SCC*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist SCC* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long SCC, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_SCC_2, replace

** Import the third part
clear
import excel using Test_esg_2.xlsx, sheet("scope3") cellrange(NTS1:UKN22) firstrow clear

gen Years = _n 
 
// Identify and convert string variables to numeric
ds SCC*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist SCC* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long SCC, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_SCC_3, replace

** Append the three parts
clear
append using Data_SCC_1 Data_SCC_2 Data_SCC_3
save Data_SCC, replace

* Save the dataset
save Data_SCC, replace

*** ----------------------Import the Return Index --------------------------- ***
** Import the first part
clear
import excel using Test_esg_2.xlsx, sheet("Total Return") cellrange(B1:GJJ22) firstrow clear
 
gen Years = _n  

// Identify and convert string variables to numeric
ds TR*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all TR variables to double
foreach var of varlist TR* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long TR, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_TR_1, replace

** Import the second part
clear
import excel using Test_esg_1.xlsx, sheet("Total Return") cellrange(GJK1:NTR22) firstrow clear
 
gen Years = _n 
 
// Identify and convert string variables to numeric
ds TR*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist TR* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long TR, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_TR_2, replace

** Import the third part
clear
import excel using Test_esg_1.xlsx, sheet("Total Return") cellrange(NTS1:UKM22) firstrow clear

gen Years = _n 
 
// Identify and convert string variables to numeric
ds TR*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all TR variables to double
foreach var of varlist TR* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long TR, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_TR_3, replace

** Append the three parts
clear
append using Data_TR_1 Data_TR_2 Data_TR_3
save Data_TR, replace

* Save the dataset
save Data_TR, replace


*** ----------------------Import the Market Capitalization --------------------------- ***
clear
import excel using Test_esg_1.xlsx, sheet("Market Capitalization") cellrange(B1:GJJ22) firstrow clear
 
gen Years = _n  

// Identify and convert string variables to numeric
ds MC*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all MC variables to double
foreach var of varlist MC* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long MC, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_MC_1, replace

** Import the second part
clear
import excel using Test_esg_1.xlsx, sheet("Market Capitalization") cellrange(GJK1:NTR22) firstrow clear
 
gen Years = _n 
 
// Identify and convert string variables to numeric
ds MC*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all SC variables to double
foreach var of varlist MC* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long MC, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_MC_2, replace

** Import the third part
clear
import excel using Test_esg_1.xlsx, sheet("Market Capitalization") cellrange(NTS1:UKM22) firstrow clear

gen Years = _n 
 
// Identify and convert string variables to numeric
ds MC*, has(type string)
foreach var of varlist `r(varlist)' {
    destring `var', replace
}

// Recast all MC variables to double
foreach var of varlist MC* {
    recast double `var'
}
 
* reshape from wide to long 
reshape long MC, i(Years)

rename _j Stock_ID
sort Stock_ID Years

save Data_MC_3, replace

** Append the three parts
clear
append using Data_MC_1 Data_MC_2 Data_MC_3
save Data_MC, replace

* Save the dataset
save Data_MC, replace


** ------------------------------------------------------------------------- **
**                  Create the Unique Dataset by Merging                     **
** ------------------------------------------------------------------------- **
clear
use Data_SC

merge m:m Stock_ID using Data_SCA.dta	
keep if _merge == 3
drop _merge

merge m:m Stock_ID using Data_SCB.dta	
keep if _merge == 3
drop _merge

merge m:m Stock_ID using Data_SCC.dta	
keep if _merge == 3
drop _merge

merge m:m Stock_ID using Data_MC.dta
keep if _merge == 3
drop _merge

merge m:m Stock_ID using Data_TR.dta
keep if _merge == 3
drop _merge


* Save the dataset
save Data, replace

*** ----------------------------------------------------------------------- ***
***               Perform a Panel Regression Analysis                       ***
*** ----------------------------------------------------------------------- ***
clear
set more off

use Data, replace

* Run a Pooled Regression
gen LSC = log(SC)

reg LSC SCA SCB SCC MC TR
reg ROE ESG LTA i.Years
reg ROE ESG LTA i.Years, robust