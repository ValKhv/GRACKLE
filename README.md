# GRACKLE
      --------------------------
The little bird-like VBA library it is named after: humble, inconspicuous, and almost useless. However, it contains basic methods to extend the capabilities of Access and Excel.
      --------------------------
      
The GRACKLE  library is a compact set of useful classes and utilities that make it easy to create complex VBA  applications. The library consists of a skeleton of related modules covering aspects of the GUI, Network, data management, extensions of the basic tools of the language.
   Using the library is possible in different ways:
-	Code snipping: borrowing individual functions and modules within your project.
-	Connect the library as an external object (via the References menu - just add the finished GRACKLE database.accdb)
-	Compile code as a dynamic link library (for example, using VBACompiler https://vbacompiler.com/ )
-	Adaptation of modules and functions to work with MS Excel in the form of Excel add-in.

## FEATURES

Structure of the library:
### THE MAIN CONTAINER - GRACKLE.ACCDB
The main container or repository is GRACKLE.accdb code (or in compiled form GRACKLE.aacde)  consists of kernel modules that form the skeleton of the library and auxiliary classes. the framework consists of the following functional modules:
  *	#_ACCESS: A library that lets you manage cons and the Access interface.  because access contains non-real components and windows as forms it requires special techniques to access their pointers.
  *	#_CHARTS: an add-on to Google Charts to overcome relatively poor graphics in Access itself. 
  *	#_COMPLEX: Inside Access, tables (the Jet kernel) may contain complex fields that hide copulas, allow you to add binary attachments, etc. The module allows you to easily operate with such fields directly from the code.
  *	#_CRYPTO: basic cryptographic primitives.
  *	#_DATETIME: functions that simplify the work with dates of different formats and their conversion among themselves.
  *	#_DIALOG: enhanced dialog boxes that allow you to use additional features when interacting with the user, such as quick choices. 
  *	#_DOCUMENTER: a set of methods for documenting code and data.
  *	#_EMAIL: adds the ability to send messages and other resources as an attachment using the Outlook COM server 
  *	#_ENVIRONMENT: allows you to get information about the environment, the operating system, including fine tuning
  *	#_EXCEL: allows you to read information or display it in Excel, operate with cells directly from the database
  *	#_EXPORT: additional opportunities to export any data, internal code, forms, queries, etc.
  *	#_FILE: file system operations, including searching for files and downloading them.
  *	#_GUI: interacting with the Windows API and directly creating and rendering forms
  *	#_HELPER: some helper functions that allow you to implement subclassing.
  *	#_JET: interaction with the SQL-core of the system, organization of breakdowns, work with DAO.
  *	#_JSON: module for working with JSON objects.
  *	#_MATH: some mathematical methods
  *	#_NET: Network communication, including managing sockets and connections, downloading remote files, and so on.
  *	#_PLUGINS: functions of working with external databases and storages, for example, connecting external forms
  *	#_STRING: string operations that complement the basic capabilities of VB
  *	#_STRUCTURES: implementation of data structures such as dictionary/hash table, etc.
  *	#_UTIL: various auxiliary tools for interacting with the clipboard, creating and configuring information objects
  *	#_WORD: interaction with MS WORD objects
 
All modules are implemented in the singleton paradigm, so that by connecting the library to your database, you can access any public functions in the modules directly or through a qualified name, for example, [#_DOCUMENTER].divider

### FORMS
   The GFORMS.accdb  forms library. This form library, when accessed by the GFORM plugin  , allows you to quickly import the form and run it directly in the client database.


## USE CASES
The Visual Basic  language started in 1991 and became widespread because of its rapid development capabilities, as it implemented the principle of connecting the programming language and the graphical interface developed by the famous programmer and technical writer Alan Cooper.  Visual Basic for Application is an offshoot of Visual Basic that is not currently supported by its manufacturer, Microsoft. VBA is built into Microsoft Office, is an interpreted language, and may also be replaced in the future with support from .Net (Visual Studio Tools for Applications).  And although the old VBA is more dead than alive, there are still a lot of applications on top of MS Office that use the capabilities of VBA.
The use of VBA has the following obvious advantages: a built-in development environment in the most widespread office suite, a COM architecture that allows you to use all the components available on the Windows  system  , easy creation of interfaces and event-oriented systems, a slender and logical syntax model.
The disadvantage of the system is the lack of support for multi-threaded applications, a somewhat outdated object model and the insufficiently broad support of the community, which is a loss compared to the popular Python.  However, there is some niche of desktop applications that make the use of VBA interesting for those who implement their own applications or automate everyday office work.
The presented library allows you to support the community by implementing a coherent set of auxiliary utilities and modules that require significant programming efforts. The tips often published on forums and discussions on how to solve certain VBA  problems look raw and need to be improved, poorly linked to each other. I hope that using the GRACKLE  library will help in this matter.
The most obvious application scenarios are:
  1.	Create your own document library. Tasks such as managing multiple files sometimes require a local tool that allows you to work offline and then sync with cloud storage such as AirTable;
  2.	Various directories of media files;
  3.	Prepare and process data by demonstrating your own interface
  4.	Prototyping applications.

## LIMITATIONS OF LIABILITIES
   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND     ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;   LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT    (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
