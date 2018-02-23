<?php

namespace Drupal\doc_serialization\Encoder;

use Drupal\Component\Serialization\Exception\InvalidDataTypeException;
use Drupal\Component\Utility\Html;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\TemplateProcessor;
use Symfony\Component\Serializer\Encoder\EncoderInterface;
use Drupal\views\Views;
use Drupal\file\Entity\File;

/**
* Adds DOCX encoder support for the Serialization API.
*/
class Docx implements EncoderInterface {

  /**
  * The format that this encoder supports.
  *
  * @var string
  */
  protected static $format = 'docx';

  /**
  * Format to write DOC files as.
  *
  * @var string
  */
  protected $docFormat = 'Word2007';

  /**
  * Constructs an DOCX encoder.
  *
  * @param string $doc_format
  *   The DOC format to use.
  */
  public function __construct($doc_format = 'Word2007') {
    $this->docFormat = $doc_format;
  }

  /**
  * {@inheritdoc}
  *
  * @throws \Drupal\Component\Serialization\Exception\InvalidDataTypeException
  */
  public function encode($data, $format, array $context = []) {
    switch (gettype($data)) {
      case 'array':
      // Nothing to do.
      break;

      case 'object':
      $data = (array) $data;
      break;

      default:
      $data = [$data];
      break;
    }

    $views_style_plugin = $context['views_style_plugin'];
    $displayHandler = $views_style_plugin->displayHandler;

    $template_fid = $displayHandler->display['display_options']['display_extenders']['doc_serialization']['doc_serialization']['template_file'][0];
    $file = File::load($template_fid);
    $file_path = \Drupal::service('file_system')->realpath($file->getFileUri());

    $templateProcessor = new TemplateProcessor($file_path);

    foreach ($data as $key => $row) {
      foreach ($row as $field => $value) {
        $templateProcessor->setValue($field, strip_tags($value));
      }
    }

    $templateProcessor->setValue('current_date', \Drupal::service('date.formatter')->format(time(), 'custom', 'd-m-Y'));

    $current_month = \Drupal::service('date.formatter')->format(time(), 'custom', 'n');
    $current_year = \Drupal::service('date.formatter')->format(time(), 'custom', 'Y');

    if ($current_month == 12) {
      $first_day_of_third_month = mktime(0, 0, 0, 0, 3, $current_year + 1);
      $first_day_of_second_month = mktime(0, 0, 0, 0, 2, $current_year + 1);
    }
    else {
      $first_day_of_third_month = mktime(0, 0, 0, $current_month + 3, 1);
      $first_day_of_second_month = mktime(0, 0, 0, $current_month + 2, 1);
    }

    $formatted_time = \Drupal::service('date.formatter')->format($first_day_of_third_month, 'custom', 'l');
    $templateProcessor->setValue('first_day_of_month', $formatted_time);

    $formatted_time = \Drupal::service('date.formatter')->format($first_day_of_second_month, 'custom', 'l');
    $templateProcessor->setValue('first_day_of_second_month', $formatted_time);

    $templateProcessor->setValue('last_year', date("Y",strtotime("-1 year")));

    ob_start();
    $templateProcessor->saveAs('php://output');
    return ob_get_clean();

    // Output escaping
    // Settings::setOutputEscapingEnabled(true);

    // try {
    //   // Instantiate a new Word object.
    //   $word = new \PhpOffice\PhpWord\PhpWord();
    //
    //   // Set the data.
    //   $this->setData($word, $data);
    //
    //   $writer = IOFactory::createWriter($word, $this->docFormat);
    //
    //   // @TODO utilize a temporary file perhaps?
    //   // @TODO This should also support batch processing.
    //   ob_start();
    //   $writer->save('php://output');
    //   return ob_get_clean();
    // }
    // catch (\Exception $e) {
    //   throw new InvalidDataTypeException($e->getMessage(), $e->getCode(), $e);
    // }
  }

  /**
  * Set document data.
  *
  * @param \PhpOffice\PhpWord\PhpWord $word
  *   The document to put the data in.
  * @param array $data
  *   The data to be put in the document.
  */
  protected function setData(PhpWord $word, array $data) {
    global $base_url;
    foreach ($data as $row) {
      $i = 0;
      // Creating a new Word section for each Views field that is displayed.
      $section = $word->addSection();
      foreach ($row as $value) {
        // $n = new HTMLtoOpenXML();
        $test = HTMLtoOpenXML::getInstance()->fromHTML($value);
        $sectionLines = $section->addTextRun();
        $sectionLines->addText($test);
        // \Drupal::logger('$value')->notice('@type', array('@type' => print_r($value, 1) ));
        // \Drupal::logger('$test')->notice('@type', array('@type' => print_r($test, 1) ));
        // @TODO No node info at this point, is there a better way then strpos?
        // if (strpos($value, '<img src="') !== FALSE) {
        //   $img_url = explode('"', explode('<img src="', $value)[1])[0];
        //   $section->addImage($base_url . $img_url);
        // }
        // else {
        //   // Parsing line breaks and paragraphs
        //   $value = $this->formatValue($value);
        //   $text = $this->getTextBetweenTags($value, 'p'); // P tags are already handled by the PHPWord processor, so we're just focussing on the BR tag here.\
        //   $sectionLines = $section->addTextRun();
        //   // \Drupal::logger('content_entity_example')->notice('@type', array('@type' => print_r($text, 1) ));
        //   foreach($text AS $line) {
        //     foreach(explode('<br>', $line) AS $v) {
        //       if ($this->startsWith($v, '<strong>') && $this->endsWith($v, '</strong>')) {
        //         $clean_text = strip_tags($v);
        //         $sectionLines->addText($clean_text, ['bold' => TRUE]);
        //       }
        //       elseif ($this->startsWith($v, '<h1>') && $this->endsWith($v, '</h1>')) {
        //         $clean_text = strip_tags($v);
        //         $sectionLines->addText($clean_text, ['bold' => TRUE, 'size' => '32px']);
        //       }
        //       elseif ($this->startsWith($v, '<h2>') && $this->endsWith($v, '</h2>')) {
        //         $clean_text = strip_tags($v);
        //         $sectionLines->addText($clean_text, ['bold' => TRUE, 'size' => '24px']);
        //       }
        //       elseif ($this->startsWith($v, '<h3>') && $this->endsWith($v, '</h3>')) {
        //         $clean_text = strip_tags($v);
        //         $sectionLines->addText($clean_text, ['bold' => TRUE, 'size' => '20.8px']);
        //       }
        //       elseif ($this->startsWith($v, '<h4>') && $this->endsWith($v, '</h4>')) {
        //         $clean_text = strip_tags($v);
        //         $sectionLines->addText($clean_text, ['bold' => TRUE, 'size' => '16px']);
        //       }
        //       else {
        //         $sectionLines->addText($v); // Adding the text.
        //       }
        //
        //       // $sectionLines->addTextBreak(); // Because there's a BR here, we want to create a newline
        //     }
        //     $sectionLines->addTextBreak(); // Because there's a BR here, we want to create a newline
        //   }
        // }
        $i++; // On to the next Views field
      }
    }
  }

  /**
  * Retrieves text from between tags
  *
  * @param string $string
  *   The string with value you want to retrieve text from
  *
  * @param string $tagname
  *   What is the tag that you want to filter? Enter without <>. For example: p
  *
  * @return matches
  *   The text within the first found tag.
  */

  private function getTextBetweenTags($string, $tagname)
  {
    $pattern = "#<\s*?$tagname\b[^>]*>(.*?)</$tagname\b[^>]*>#s";
    preg_match_all($pattern, $string, $matches);
    return $matches[1];
  }


  function startsWith($haystack, $needle)
  {
    $length = strlen($needle);
    return (substr(trim($haystack), 0, $length) === $needle);
  }

  function endsWith($haystack, $needle)
  {
    $length = strlen($needle);

    return $length === 0 ||
    (substr(trim($haystack), -$length) === $needle);
  }

  /**
  * Formats a single value for a given value.
  *
  * @param string $value
  *   The raw value to be formatted.
  *
  * @return string
  *   The formatted value.
  */
  protected function formatValue($value) {
    $value = Html::decodeEntities($value);
    $value = strip_tags($value, '<p><br><strong><h1><h2><h3><h4><italic><u>');

    return $value;
  }

  /**
  * {@inheritdoc}
  */
  public function supportsEncoding($format) {
    return $format === static::$format;
  }

}
