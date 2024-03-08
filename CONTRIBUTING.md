
## How to contribute to NanoXLSX

### Preamble
Thank yo very much for your interest in NanoXLSX. This library is maintained completely based on the community and is not backed by a company or organization. Any contribution is highly appreciated and helps to increase the quality and relevance of NanoXLSX.

However, to ensure a good experience for everyone (library users, developers, or other contributors), we kindly ask to follow the recommendations in this document when doing a contribution to the library

### Creating an issue

#### General
* Please use one of the provided issue templates if applicable
* Please use English as common language. If you are not confident, no worries. [DeepL]( https://www.deepl.com/translator) or other online translators can help to write an issue

#### Reporting a bug

* When reporting a bug, please provide as much information as feasible. You can follow the requirement in the template
* Please provide always the used version of NanoXLSX and at least the used environment
  * .NET framework version
  * Operating system
  * Used IDE if applicable
* Please describe as exact as possible how to reproduce the bug
* Attach a small demo Excel file in the issue if the bug relates to a file operation, like reading or writing data
* For your own safety, please do not upload Excel files that contains any real business data about employees, customers or other data that could lead to a misuse of such data by a malicious 3rd party
  
### Creating a Pull Request (PR)

* The best way to start a PR process is to open an issue first. Create the PR and link the PR in the issue. In this issue, the topic of the PR can be described and discussed
* If no issue is created, please provide exact information in the PR header, what the changes are supposed to achieve
* Please set the branch **dev-pr** as base of your PR. This is a branch solely designated to PRs and is used to checkout and test the changes before merging them to dev and later to master
  
### PRs that probably cannot be accepted or needs rework

* No description what the PR is supposed to achieve (add at least a message that a bug was found, and this PR fixes it, if applicable)
* PR is only applying code formatting or doing refactoring without any functional changes (please consider opening an issue instead to discuss code styles or formatting)
* PR breaks existing unit tests (please fix the tests if the original implementation leads to a wrong result) 
* PR removes public functionality (part of the API)
* PR alters the behavior of a public API function radically, without addressing a bug (please open an issue first to discuss the options about a possible broken API function)
* PR adds proprietary functionality that is not compatible with the OOXML standard (XLSX)
* PR solves a particular problem at one specific location of the library but does not cover other instances that have the exact same issue (may be addressed as PR comment)
* PR introduces an external  NuGet dependency or assumes the availability of a specific system library or resource
* Obscure / *unclean* PRs:
  * Provides an unclear function or use of .NET framework without explanation
  * Uses another code style as the rest of the library (e.g. suddenly using snake case for variables)
  * Contains non-English variable or function names, or comments
  * Contains code that is commented out
  * Excessive use of the var keyword (The goal of the library is to make its usage as clear as possible and that snippets are always appropriately typed)
  * Uses cryptic variable or function names (`CalcAmount` may be OK as function name , but `GT56R` may not)
  * Introduces (inline) hard-coded values that are either already defined as constants or could be substituted by an existing enum or similar code parts
  * Code that is clearly copied from somewhere else or is AI generated
    * External sources: Ensure compliance with the licenses. Don’t use code that may have a more restrictive license than NanoXLSX (MIT)
    * AI: That’s great! But please, revise the code, check it, and remove unnecessary comments. Ensure that there could not arise a licensing issue  
