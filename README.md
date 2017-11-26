# Word Serialization

This module provides an Word encoder for the Drupal 8 Serialization API. This
enables the DOC(X) format to be used for data output (and potentially input,
eventually). For example:

  * Views can output DOC(X) data via a 'Word Export' display in a View.
  * Module developers can leverage DOC(X) as a format when using the 
    Serialization API.

#### Installation

  * Download and install
    [PHPOffice/PHPWord](https://github.com/PHPOffice/PHPWord).
    and all of it's dependencies:
    * [zendframework/zend-escaper 2.4.*](https://github.com/zendframework/zend-escaper/tree/release-2.4.13)
    * [zendframework/zend-stdlib 2.4.*](https://github.com/zendframework/zend-stdlib/tree/release-2.4.13)
    * [zendframework/zend-validator 2.4.*](https://github.com/zendframework/zend-validator/tree/release-2.4.13)
    * [zendframework/zend-stdlib 2.4.*](https://github.com/zendframework/zend-stdlib/tree/release-2.4.13)
    * [phpoffice/common 0.2.6](https://github.com/PHPOffice/Common/tree/0.2.6)
    * [pclzip/pclzip": ^2.8](https://github.com/ivanlanin/pclzip/tree/2.8.2) 
    
    The preferred installation method is to 
    [use Composer](https://www.drupal.org/node/2404989).
  * Enable the `doc_serialization` module.

#### Creating a view with a DOC display

  1. Create a new view
  2. Add a *Word Export* display.
  3. Select either 'docx' or 'doc' for the accepted request formats under
     `Format -> Word export -> Settings`.
  4. Add desired fields to the view.
  5. Add a path, and optionally, a filename pattern.
