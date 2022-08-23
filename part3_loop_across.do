// Load data from github repo
use "https://github.com/andrewmcamp/Stata-to-PPT-Example/raw/main/gapminder.dta", clear

// ======================= Part 3: Loop Across Continents =========================================
levelsof continent, local(continents)
foreach c in `continents' {
	// Get the label for each level of the continent variable to be used when saving the file
	local cname: label continent `c'
	di "Creating .ppt for `cname'."

	// Preserving/restoring the data before we limit to each continent
	preserve
		keep if continent == `c'
	
		// First, clean up any files in the current directory leftover from previous loop
		cap erase "template.zip"
		cap shell rd "_rels" /s /q
		cap shell rd  "docProps" /s /q
		cap shell rd  "ppt" /s /q
		cap erase "[Content_Types].xml"

		// Next copy our template and rename as a .zip file
		shell copy "template.pptx" "template.zip"

		// Next, unzip our template
		unzipfile "template.zip"

		// Now we can export our graphs using the same code as from part 1, but different paths

			// Graph 1: GDP Per Capita vs. Life Expectancy
			twoway scatter gdpPercap lifeExp
			graph export "ppt/media/image2.svg", replace width(13in) height(7in)
			
			// Graph 2: GDP Per Capita vs. Life Expectancy by Year
			twoway scatter gdpPercap lifeExp, by(year)
			graph export "ppt/media/image4.svg", replace width(13in) height(7in)

			// Graph 3: Bar chart for populations over 1,000,000 in 2007
			graph hbar pop if year == 2007 & pop > 1000000, over(country, sort(pop))
			graph export "ppt/media/image6.svg", replace width(13in) height(7in)

			// Table 1: Average Life Expectancy, Population, and GDP Per Capita by Country
			// Getting values to put in using tabstat/tabstatmat
			tabstat lifeExp, by(year) statistics(mean sd min max) nosep save
			tabstatmat table1, nototal

			// 	Using putexcel to export the summary statistics above
			putexcel set "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx", modify
			putexcel A2 = matrix(table1), names

		// To finish this iteration of the loop, we simply zip the folder back up
		zipfile _rels docProps ppt "[Content_Types].xml", saving("report_`cname'.pptx", replace)

		// And finally we clean up any of the intermediary files created during the loop.
		cap erase "template.zip"
		cap shell rd "_rels" /s /q
		cap shell rd  "docProps" /s /q
		cap shell rd  "ppt" /s /q
		cap erase "[Content_Types].xml"
		
restore

}
