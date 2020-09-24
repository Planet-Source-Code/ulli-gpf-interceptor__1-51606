GPF-Interceptor (based on a previous submission to PSC, tnx to the original author)
---------------

This project shows how you can intercept GPFs. Instead of crashing your application in case of a GPF, a PopUp is openend showing error details and giving the user the choice between continuing the buggy application or terminating it in a neat and proper way.

To use the GPF interceptor simply add cException, fException, mException, rException, and cTooltips to your project. See fTest how to activate and de-activate the GPF interceptor.