<?php
/**
 * \PhpOffice\PhpSpreadsheet\Spreadsheet
 *
 * Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */

\PhpOffice\PhpSpreadsheet\Spreadsheet_Autoloader::Register();
//    As we always try to run the autoloader before anything else, we can use it to do a few
//        simple checks and initialisations
//\PhpOffice\PhpSpreadsheet\Spreadsheet_Shared_ZipStreamWrapper::register();
// check mbstring.func_overload
if (ini_get('mbstring.func_overload') & 2) {
    throw new \PhpOffice\PhpSpreadsheet\Exception('Multibyte function overloading in PHP must be disabled for string functions (2).');
}
\PhpOffice\PhpSpreadsheet\Shared\StringHelper::buildCharacterSets();


/**
 * \PhpOffice\PhpSpreadsheet\Spreadsheet_Autoloader
 *
 * @category    \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package     \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @copyright   Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Spreadsheet_Autoloader
{
    /**
     * Register the Autoloader with SPL
     *
     */
    public static function Register() {
        if (function_exists('__autoload')) {
            //    Register any existing autoloader function with SPL, so we don't get any clashes
            spl_autoload_register('__autoload');
        }
        //    Register ourselves with SPL
        if (version_compare(PHP_VERSION, '5.3.0') >= 0) {
            return spl_autoload_register(array('\PhpOffice\PhpSpreadsheet\Spreadsheet_Autoloader', 'Load'), true, true);
        } else {
            return spl_autoload_register(array('\PhpOffice\PhpSpreadsheet\Spreadsheet_Autoloader', 'Load'));
        }
    }   //    function Register()


    /**
     * Autoload a class identified by name
     *
     * @param    string    $pClassName        Name of the object to load
     */
    public static function Load($pClassName){
        if ((class_exists($pClassName,FALSE)) || (strpos($pClassName, '\PhpOffice\PhpSpreadsheet\Spreadsheet') !== 0)) {
            //    Either already loaded, or not a \PhpOffice\PhpSpreadsheet\Spreadsheet class request
            return FALSE;
        }

        $pClassFilePath = PHPEXCEL_ROOT .
                          str_replace('_',DIRECTORY_SEPARATOR,$pClassName) .
                          '.php';

        if ((file_exists($pClassFilePath) === FALSE) || (is_readable($pClassFilePath) === FALSE)) {
            //    Can't load
            return FALSE;
        }

        require($pClassFilePath);
    }   //    function Load()

}
