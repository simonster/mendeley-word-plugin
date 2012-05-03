# The Mendeley Word for Windows Plugin

This is the Mendeley Word for Windows plugin code, extracted from the template that they ship with
Word as of version 1.5.1 and distributed under the terms of the Educational Community License. It is
based on the Visual Basic code I wrote 5 years ago for [Zotero](http://www.zotero.org/).

In 2009, I rewrote
the [Zotero Word for Windows plugin](http://www.zotero.org/support/word_processor_plugin_installation_for_zotero_2.1) in
[C++](https://github.com/zotero/zotero-word-for-windows-integration). That code is more versatile,
and is probably a better starting point for any free (as in AGPL) reference mangement software
looking to implement integration with Word. However, I provide this code in case non-free
implementers are interested, and because it's partly mine and it wants to be free.

Do not ask me for support or assistance. I will never touch this code again.

This distribution includes the following modules and class modules from Mendeley-1.5.1.dotm:

* Mendeley
* MendeleyLib
* ZoteroLib
* EventClassModule

It does not contain the following modules, which lack an ECL copyright notice and, in the absence of
further information from Mendeley Ltd., are assumed to be closed source.

* MendeleyDataTypes
* MendeleyRibbon
* MendeleyUnitTests
* StyleListModel

Compatible versions of some of these files may be available in Mendeley's
[openoffice-plugin](https://github.com/Mendeley/openoffice-plugin/tree/master/src) repository.

	Copyright (c) 2009-2012 Mendeley Ltd.
	Copyright (c) 2006      Center for History and New Media
	                        George Mason University, Fairfax, Virginia, USA
	                        http://chnm.gmu.edu