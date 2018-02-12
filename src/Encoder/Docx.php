<?php

namespace Drupal\doc_serialization\Encoder;

use Drupal\Component\Serialization\Exception\InvalidDataTypeException;
use Drupal\Component\Utility\Html;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use Symfony\Component\Serializer\Encoder\EncoderInterface;

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

    try {
      // Instantiate a new Word object.
      $word = new \PhpOffice\PhpWord\PhpWord();

      // Set the data.
      $this->setData($word, $data);

      $writer = IOFactory::createWriter($word, $this->docFormat);

      // @TODO utilize a temporary file perhaps?
      // @TODO This should also support batch processing.
      ob_start();
      $writer->save('php://output');
      return ob_get_clean();
    }
    catch (\Exception $e) {
      throw new InvalidDataTypeException($e->getMessage(), $e->getCode(), $e);
    }
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
        // @TODO No node info at this point, is there a better way then strpos?
        if (strpos($value, '<img src="') !== FALSE) {
          $img_url = explode('"', explode('<img src="', $value)[1])[0];
          $section->addImage($base_url . $img_url);
        }
        else {
			// Parsing line breaks and paragraphs
			$value = $this->formatValue($value);
			$text = $this->getTextBetweenTags($value, 'p'); // P tags are already handled by the PHPWord processor, so we're just focussing on the BR tag here.
			$sectionLines = $section->addTextRun();
			foreach($text AS $line) {
			  foreach(explode('<br>', $line) AS $v) {
			    $sectionLines->addText($v); // Adding the text.
				$sectionLines->addTextBreak(); // Because there's a BR here, we want to create a newline
			  }
			}
        }
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
    $value = strip_tags($value, '<p><br>');
	
    return $value;
  }

  /**
   * {@inheritdoc}
   */
  public function supportsEncoding($format) {
    return $format === static::$format;
  }

}

