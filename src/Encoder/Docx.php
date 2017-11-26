<?php

namespace Drupal\doc_serialization\Encoder;

/**
 * Adds DOCX encoder support for the Serialization API.
 */
class Docx extends Doc {

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
   * Constructs an DOC encoder.
   *
   * @param string $doc_format
   *   The DOC format to use.
   */
  public function __construct($doc_format = 'Word2007') {
    $this->docFormat = $doc_format;
  }

  /**
   * {@inheritdoc}
   */
  protected function setSettings(array $settings) {}

}
