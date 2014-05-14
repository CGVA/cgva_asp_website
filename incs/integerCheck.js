<script language="JavaScript">

<!--

function isDigit (c)

{

            return ((c >= "0") && (c <= "9"))

}

 

function isInteger(s)

{

            var i;

            // Search through string's characters one by one

            // until we find a non-numeric character.

            // When we do, return false; if we don't, return true.

 

            for (i = 0; i < s.length; i++)

            {

             // Check that current character is number.

             var c = s.charAt(i);

 

             if (!isDigit(c)) return false;

            }

 

            // All characters are numbers.

            return true;

}

 

function isDecimal(s)

{

            var decimalCount = 0;

            var i;

            // Search through string's characters one by one

            // until we find a non-decimal character.

            // When we do, return false; if we don't, return true.

 

            for (i = 0; i < s.length; i++)

            {

                        // Check that current character is number.

                        var c = s.charAt(i);

 

                        //alert(c);

                        if(!isDigit(c) && c != ".")

                        {

                                    return false;

                        }

 

                        if(c == ".")

                        {

                                    //alert("here");

                                    decimalCount += decimalCount + 1;

                                    if(decimalCount > 1)

                                    {

                                                return false;

                                    }

                        }

                        //else

                        //{

                        //          alert("oops");

                        //}

            }

 

            // All characters are numbers.

            return true;

}
-->
</script>