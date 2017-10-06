# connectivity-support
[![Community Supported](https://img.shields.io/badge/Support%20Level-Community%20Supported-457387.svg)](https://www.tableau.com/support-levels-it-and-developer-tools)

Provides a set of tools for diagnosing common infrastructure concerns in the connectivity layer.

## SAPBW

### bw-sso-test-tool
This is a Windows console application that looks up SAP BW-related configuration information and attempts to connect to a specified BW instance, producing output that can help diagnose configuration issues. It captures relevant configuration settings, gives warnings when it finds settings that are missing or different than expected, and produces trace files for the connection attempt. It is useful for both kerberos-based (think Tableau Desktop to BW) and delegation via Server Side Trust (think Viewer Credentials authentication in Tableau Server). 

#### Is bw-sso-test-tool supported?

The bw-sso-test-tool is made available under community support. Despite efforts to write good and useful code there may be bugs that cause unexpected and undesirable behavior. The software is strictly “use at your own risk.”

The good news: This is intended to be a self-service tool. You are free to modify it in any way to meet your needs.
