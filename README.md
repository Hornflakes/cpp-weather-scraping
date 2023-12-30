# fast-weather-scraping

This repo contains a weather web scraping C++ script (code and .bat) I made for my dad.

He records and analyzes weather data for our orchard. 🍇🌳🍐🍑🍎🍒

[Original Python version repo 🐍](https://github.com/Hornflakes/weather-scraping)

This C++ version is also an attempt at adopting (dys)functional programming principles.

## What does it do?

-   extracts weather data from [freemeteo.ro](https://freemeteo.ro/vremea)
-   exports the data into Excel

## Setup

Make sure the excel files are in the same folder from which you run the script (be it .exe or .bat).

**Dev:**

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; I used MinGW64 and MSYS2.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; **Install packages:** `curl` `gumbo` `xlnt`

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; **Compile to .exe:** `g++ -o weather_script main.cc -lcurl -lgumbo -lxlnt`

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; or use your preferred build system

**User:**

1. download `Weather_Scraping.zip` from **Releases**

    or copy what's inside the `dist` folder

2. replace `weather_data.xlsx` with your own excel

    or use it (skip the next 2 steps)

3. add the date (dd.mm.yyyy) after which you want to get data
4. update `config.xlsx`

    - your excel name
    - your sheet name
    - the letter(s) header of the first (date) column

5. run `weather_script.bat` and wait for the console window to close
6. PROFIT 📈📈📈

##

The CA certificate for the data request can expire. In that case, you can replace it.

The certificate is in the `certs` folder. [You can download certificates here.](https://curl.se/docs/caextract.html)
