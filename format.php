<?php
// This file is part of Moodle - http://moodle.org/
//
// Moodle is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// Moodle is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with Moodle.  If not, see <http://www.gnu.org/licenses/>.

/**
 * CSV format question importer.
 *
 * @package    qformat_csv
 * @copyright  2021 Gopal Sharma <gopalsharma66@gmail.com>
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */


defined('MOODLE_INTERNAL') || die();

/*
 CSV format - a simple format for creating multiple and single choice questions.
 * The format looks like this for simple csv file with minimum columns:
 * questionname, questiontext, A,   B,   C,   D,   Answer 1,    Answer 2
 * Question1, "3, 4, 7, 8, 11, 12, ... What number should come next?",7,10,14,15,D
 *
 *
 * The format looks like this for Extended csv file with extra columns columns:
 * questionname, questiontext, A,   B,   C,   D,   Answer 1,    Answer 2,
   answernumbering, correctfeedback, partiallycorrectfeedback, incorrectfeedback, defaultmark
 * Question1, "3, 4, 7, 8, 11, 12, ... What number should come next?",7,10,14,15,D, ,
   123, Your answer is correct., Your answer is partially correct., Your answer is incorrect., 1
 *
 *
 *  That is,
 *  + first line contains the headers separated with commas
 *  + Next line contains the details of question, each line contain
 *  one question text, four option, and either one or two answers again all separated by commas.
 *  Each line contains all the details regarding the one question ie. question text, options and answer.
 *  You can also download the sample file for your reference.
 *
 * @copyright 2018 Gopal Sharma <gopalsharma66@gmail.com>
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */

$globals['header'] = true;
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class qformat_csv extends qformat_default {

    public function provide_import() {
        return true;
    }

    public function provide_export() {
        return true;
    }

    /**
     * @return string the file extension (including .) that is normally used for
     * files handled by this plugin.
     */
    public function export_file_extension() {
        return '.xlsx';
    }

    /**
     * Return complete file within an array, one item per line
     * @param string filename name of file
     * @return mixed contents array or false on failure
     */
    protected function readdata($filename) {
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load($filename);
        $worksheet = $spreadsheet->getActiveSheet();
        $lines = [];
        for ($i = 2; $i <= 1454; $i += 4) {
            $lines[] = [
                'name' => $worksheet->getCell('A' . $i)->getValue(),
                'questiontext' => $worksheet->getCell('E' . $i)->getValue() . '<br/>' . $worksheet->getCell('I' . $i)->getValue(),
                'answer' => [
                    $this->text_field($worksheet->getCell('F' . $i)->getValue()),
                    $this->text_field($worksheet->getCell('F' . (string)($i + 1))->getValue()),
                    $this->text_field($worksheet->getCell('F' . (string)($i + 2))->getValue()),
                    $this->text_field($worksheet->getCell('F' . (string)($i + 3))->getValue()),
                ],
                'fraction' => [
                    $worksheet->getCell('G' . $i)->getValue() == 1 ? 1 : 0,
                    $worksheet->getCell('G' . $i)->getValue() == 2 ? 1 : 0,
                    $worksheet->getCell('G' . $i)->getValue() == 3 ? 1 : 0,
                    $worksheet->getCell('G' . $i)->getValue() == 4 ? 1 : 0,
                ],
                'feedback' => [
                    $this->text_field(''),
                    $this->text_field(''),
                    $this->text_field(''),
                    $this->text_field('')
                ],
                'generalfeedback' => $worksheet->getCell('H' . $i)->getValue() . '<br/>' . $worksheet->getCell('J' . $i)->getValue(),
                'category' => trim($worksheet->getCell('B' . $i)->getValue())
            ];
        }
        return $lines;
    }

    public function readquestions($lines) {
        global $CFG, $DB;
        question_bank::get_qtype('multichoice'); // Ensure the multianswer code is loaded.
        $questions = array();
        foreach ($lines as $line) {
            $catquestion = $this->defaultquestion();
            $catquestion->category = 'top/' . mb_strtolower($line['category']);
            $catquestion->qtype = 'category';
            $questions[] = $catquestion;
            $question = $this->defaultquestion();
            $question->qtype = 'multichoice';
            $question->answernumbering = '123';
            $question->single = 1;
            $question->name = $line['name'];
            $question->questiontext = html_entity_decode($line['questiontext']);
            $question->questiontextformat = 1;
            $question->answer = $line['answer'];
            $question->fraction = $line['fraction'];
            $question->feedback = $line['feedback'];
            $question->generalfeedback = html_entity_decode($line['generalfeedback']);
            $question->generalfeedbackformat = FORMAT_HTML;
            $questions[] = $question;
        }
        return $questions;
    }
    protected function text_field($text) {
        return array(
            'text' => html_entity_decode(trim($text)),
            'format' => FORMAT_HTML,
            'files' => array(),
        );
    }

    public function readquestion($lines) {
        // This is no longer needed but might still be called by default.php.
        return;
    }

    public function writequestion($question) {
        global $OUTPUT;
        $expout = "";
        $rightanswer = "";
        $answercount = 0;
        $rightanswercount = 0;
        // Output depends on question type.
        // CSV Header should be printed only once.
        if ($globals['header']) {
                $expout .= "questionname,questiontext,A,B,C,D,Answer 1,Answer 2,";
                $expout .= "answernumbering, correctfeedback, partiallycorrectfeedback, incorrectfeedback, defaultmark";
                $globals['header'] = false;
        }

        switch($question->qtype) {
            case 'multichoice':
                if (count($question->options->answers) != 4 ) {
                    break;
                }
                $expout .= '"'.$question->name.'"'.',';
                $expout .= '"'.$question->questiontext.'"'.',';
                foreach ($question->options->answers as $answer) {
                    $answercount++;
                    if ($answer->fraction == 1 && $question->options->single) {
                        switch ($answercount) {
                            case 1:
                                $rightanswer = 'A'.', ,';
                                break;
                            case 2:
                                $rightanswer = 'B'.', ,';
                                break;
                            case 3:
                                $rightanswer = 'C'.', ,';
                                break;
                            case 4:
                                $rightanswer = 'D'.', ,';
                                break;
                            default:
                                $rightanswer = '';
                                break;
                        }
                    } else if ($answer->fraction == 0.5 && !$question->options->single) {
                        $rightanswercount ++;
                        $comma = ",";
                        if ( $rightanswercount <= 1 ) {
                            $comma = ","; // Add comma  to first answer i.e. to 'Answer 1'.
                        }
                        switch ($answercount) {
                            case 1:
                                $rightanswer .= 'A'.$comma;
                                break;
                            case 2:
                                $rightanswer .= 'B'.$comma;
                                break;
                            case 3:
                                $rightanswer .= 'C'.$comma;
                                break;
                            case 4:
                                $rightanswer .= 'D'.$comma;
                                break;
                            default:
                                $rightanswer = '';
                                break;
                        }

                    }
                    $expout .= '"'.$answer->answer.'"'.',';
                }
                $expout .= $rightanswer;
                $expout .= '"'.$question->options->answernumbering.'"'.',';
                $expout .= '"'.$question->options->correctfeedback.'"'.',';
                $expout .= '"'.$question->options->partiallycorrectfeedback.'"'.',';
                $expout .= '"'.$question->options->incorrectfeedback.'"'.',';
                $expout .= '"'.$question->defaultmark.'"'.',';

            break;
        }
        return $expout;
    }
}
