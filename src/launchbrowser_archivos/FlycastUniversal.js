// Copyright 1999 Flycast Communications. All rights reserved.
// Patent Pending
// Version 3.5.3

FlycastLoaded		= true;
FlycastRandom		= 0;
FlycastFoundMSIE	= navigator.userAgent.indexOf("MSIE") >= 0;
FlycastFoundMSIE2	= navigator.userAgent.indexOf("MSIE2") >= 0 || navigator.userAgent.indexOf("MSIE 2") >= 0;
FlycastFoundMSIE3	= navigator.userAgent.indexOf("MSIE 3") >= 0;
FlycastFoundNN		= navigator.userAgent.indexOf("Mozilla/") >= 0 && !FlycastFoundMSIE;
FlycastFoundNN2		= navigator.userAgent.indexOf("Mozilla/2.") >= 0 && !FlycastFoundMSIE;
FlycastFoundNN3		= navigator.userAgent.indexOf("Mozilla/3.") >= 0 && !FlycastFoundMSIE;

function FlycastDeliverAd () {

	FlycastAdServer		= "http://adex3.flycast.com/server";

	if (FlycastNewAd) {
		FlycastNow		= new Date();
		FlycastRandom	= FlycastNow.getTime();
		FlycastRandom	= FlycastRandom % 10000000;
		if (!(FlycastFoundNN2 || FlycastFoundMSIE3))
			FlycastRandom	+= Math.floor(Math.random() * 100);
	}

	FlycastSiteInfo		= FlycastSite + "/" + FlycastPage + "/" + FlycastRandom;

	if (FlycastFoundMSIE2) {
		document.write('<A HREF="' + FlycastAdServer + '/click/' +  FlycastSiteInfo + '"><IMG SRC="' + FlycastAdServer + '/ad/' +  FlycastSiteInfo + '" scrolling="no" marginwidth=0 marginheight=0 frameborder=0 vspace=0 hspace=0 width=' + FlycastWidth + ' height=' + FlycastHeight + '></A>');
	}
	else if (FlycastFoundMSIE)  {
		document.write('<IFRAME SRC="' + FlycastAdServer + '/iframe/' +  FlycastSiteInfo + '" scrolling="no" marginwidth=0 marginheight=0 frameborder=0 vspace=0 hspace=0 width=' + FlycastWidth + ' height=' + FlycastHeight + '></IFRAME>');
	}
	else {
		document.write('<S' + 'CRIPT SRC="' + FlycastAdServer + '/js/' +  FlycastSiteInfo + '" LANGUAGE=JAVASCRIPT></SCRIPT>');
	}
}

if (FlycastFoundNN3 && FlycastPrintTag) {
	FlycastLoaded	= false;
	FlycastDeliverAd();
}
