---
layout: post
title:  "Export Distribution Groups and their Members from Microsoft 365 to CSV"
date:   2022-10-08 09:00:00 +0100
categories: microsoft 365, powershell
---

# Introduction
I always find myself wanting to export data and present it in a meaningful and easy way. Whether this is to present to a client or to discuss internally - the aim is always to simplify the data so that it can be actioned.

One simple example of this was a client wanting to "tidy up" their Microsoft 365 Distribution Lists. They wanted an export of all of the users which were in the groups. Sounds simple right?

Well, it is. However, one piece of code I keep coming back to is when you have one "parent" key, such as a group, and a number of values underneath it. The ultimate goal is to list the groups as follows:

![Group Members Excel Example](/_site/assets/group_members_example_excel_jpg.jpg)

To do this, we need a nested `"foreach"` statement within `PowerShell`. I'll breakdown the code below as follows, which can also be found on my [GitHub][pythonshell-gh].

# Pre-requisites
To get started, you will need to install the `Exchange Online` module from `Microsoft`: [Exchange Online Module](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/)

# Getting Started
The first thing we will do is connect to Exchange Online of the tenant we wish to export the information from and create a blank array that we will add all of our juicy data to.

{% highlight Powershell %}
# Log in using your Global/Exchange admin username and password - supports MFA
Connect-ExchangeOnline
# Initialise a blank array for you to add the groups and members to.
$group_list = @()
{% endhighlight %}


[pythonshell-gh]:  https://github.com/pythonshell-io/ms-groups-members