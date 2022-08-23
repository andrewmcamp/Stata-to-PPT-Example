// Load data from github repo
use "https://github.com/andrewmcamp/Stata-to-PPT-Example/raw/main/gapminder.dta", clear

// ======================= Part 1: Make Example Graphs ============================================
// For example graphs/tables, we'll just use the Americas
keep if continent == 2

// Graph 1: GDP Per Capita vs. Life Expectancy
twoway scatter gdpPercap lifeExp
graph export "graph1.svg", replace width(13in) height(7in)

// Graph 2: GDP Per Capita vs. Life Expectancy by Year
twoway scatter gdpPercap lifeExp, by(year)
graph export "graph2.svg", replace width(13in) height(7in)

// Graph 3: Bar chart for population in 2007
graph hbar pop if year == 2007, over(country, sort(pop))
graph export "graph3.svg", replace width(13in) height(7in)

// Table 1: Average Life Expectancy, Population, and GDP Per Capita by Country
// Getting values to put in using tabstat/tabstatmat
tabstat lifeExp, by(year) statistics(mean sd min max) nosep save
tabstatmat table1, nototal

// 	Using putexcel to export the summary statistics above
putexcel set table1, replace
putexcel A2 = matrix(table1), names

// The next step is to create a template powerpoint to base our continent-specific reports on.
