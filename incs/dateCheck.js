<script language="JavaScript">

<!--
function checkDate(dateStr,dateType) {

 

            var datePat = /^(\d{1,2})(\/|-)(\d{1,2})\2(\d{4})$/;

            var matchArray = dateStr.match(datePat); // is the format ok?

 

            if (matchArray == null) {

                        alert("The " + dateType + " date " + dateStr + " is not a valid format.")

                        return false;

            }

            month = matchArray[1]; // parse date into variables
            day = matchArray[3];
            year = matchArray[4];

            if (month < 1 || month > 12) { // check month range

                        alert("The " + dateType + " Month must be between 1 and 12.");

                        return false;

            }

 

            if (day < 1 || day > 31) {

                        alert("The " + dateType + " Day must be between 1 and 31.");

                        return false;

            }

 

            if ((month==4 || month==6 || month==9 || month==11) && day==31) {

                        alert("The " + dateType + " Month " + month + " doesn't have 31 days.")

                        return false;

            }

 

            if (month == 2) { // check for february 29th

                        var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));

                        if (day>29 || (day==29 && !isleap)) {

                                    alert("February " + year + " doesn't have " + day + " days.");

                                    return false;

                        }

            }

 

            return true;  // date is valid

}
-->
</script>
