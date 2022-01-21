# IsIt

IsIt compares the current date or time using the comparison operator passed by /compare to the respective value passed by /value. Depending on the result of the comparison, it sets the _ERRORLEVEL_ environment variable on exit, where

- ERRORLEVEL = 0  
The comparison failed/didn't match.
- ERRORLEVEL = 1  
The comparison succeeded/matched.
- ERRORLEVEL = 254  
One or more invalid parameters were passed.
- ERRORLEVEL = 255  
An error occurred during program execution.

---

## Usage

`IsIt /compare=<comparison operator> /unit=<unit/part of date/time to compare> /value=<compare current date or time against this value> [/firstdayofweek=<first day of week>]`

   or  

`IsIt /c=<comparison operator> /u=<unit/part of date/time to compare> /v=<compare current date or time against this value> [/fdw=<first day of week>]`

e.g.

- `IsIt /c=-2 /u=date /v=20220321`  
Evaluates if the current date is before _(/c=-2)_ 21th March 2022 _(/v=20220321)_.

- `IsIt /c=1 /u=time /v=233000`  
Evaluates if the current time is exact or after _(/c=1)_ half to midnight _(/v=233000)_.

## Parameters

- `/c` or `/compare`  
  Comparison operations to perform. Valid values are:  
  - -2 = Before  
  - -1 = Before or exact  
  - 0 = Exact  
  - 1 = After or exact  
  - 2 = After

- `/u` or `/unit`  
  Unit/part of the current date or time which should be compared to the value passed. Valid values are:  
  **Date**  
  - _date_ = Full date (year, month and day)  
  - _y_ or _year_ = Year only  
  - _m_ or _month_ = Month only  
  - _d_ or _day_ = Day only  
  - _ym_ or _yearmonth_ = Year and month  
  - _md_ or _monthday_ = Month and day  
  - _wd_ or _weekday_ = Day of week _(Su-Sa)_  

  **Time**  
  - _time_ = Full time _(hour, minute and second)_  
  - _h_ or _hour_ = Hour only  
  - _n_ or _minute_ = Minute only _('n' is the odd outlier!)_  
  - _s_ or _second_ = Second only  
  - _hm_ or _hourminute_ = Hour and minute

- `/v` or `/value`  
  Value against which the current date/time should be compared. The format for passing values obviously depends on the unit passed _(see above)_. Valid value formats are:  
  **Date**  
  - for `/u=date`  
  _yyyymmdd_  
  - for `/u=y`  
  _yyyy_  
  - for `/u=m`  
  _mm_  
  - for `/u=d`  
  _dd_  
  - for `/u=ym`  
  _yyyymm_  
  - for `/u=md`  
  _mmdd_  
  - for `/u=wd`  
  _0 (Sunday)_ to _6 (Saturday)_  

  **Time**  
  - for `/u=time`  
  _hhnnss_  
  - for `/u=h`  
  _hh_  
  - `/u=n`  
  _nn_  
  - for `/u=s`  
  _ss_  
  - for `/u=hn`  
  _hhnn_

  Where:  
  - _y_ = Year  
  - _m_ = Month  
  - _d_ = Day  
  - _h_ = Hour _(format 24 HH)_  
  - _n_ = Minute  
  - _s_ = Second

`/fdw` or `/firstdayofweek`  
 _(Optional)_ Set which day is considered to be the first day of the week for _/unit=wd_. Valid values are 0 _(Sunday)_ to 6 _(Saturday)_. If omitted, 1 _(Monday)_ is the default value.

## Creating a log file

IsIt writes all output to STDOUT. So in order to produce a log file of its actions, simply redirect the output to a file via _'> log\_file\_name'_.
