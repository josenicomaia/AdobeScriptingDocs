// Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.

// on localized builds we pull the $$$/Strings from a .dat file, see documentation for more details
$.localize = true;

main(arguments[0]);

function main()
{
	try{
		var myArg = arguments[0];
		var strLocString = localize( myArg );
		return strLocString;

	} catch (myError) {
		alert(myError);
	}

}