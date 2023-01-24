import unittest  

from partname_module.name_changer import make_part_names_list
from cell import Cell


class TestPartName(unittest.TestCase): 

    def test_make_part_names_list(self):

        sheet_with_3_pos = {
                        'G1': Cell('720000'),
                        'A{}'.format(9): Cell('K00'),
                        'B{}'.format(9): Cell('0582'),
                        'D{}'.format(9): Cell('3'),
                        'E{}'.format(9): Cell('3'),
                        'H{}'.format(9): Cell('3.3547'),
                        'L{}'.format(9): Cell('T=3,3'),
                        'C{}'.format(9): Cell('A'),
                        'A{}'.format(10): Cell('K00-01'),
                        'B{}'.format(10): Cell('0202'),
                        'D{}'.format(10): Cell('1'),
                        'E{}'.format(10): Cell(''),
                        'H{}'.format(10): Cell('3.3547'),
                        'L{}'.format(10): Cell(''),
                        'C{}'.format(10): Cell('A'),
                        'A{}'.format(11): Cell(None),
                        'B{}'.format(11): Cell(None),
                        'D{}'.format(11): Cell(None),
                        'E{}'.format(11): Cell(None),
                        'H{}'.format(11): Cell(None),
                        'L{}'.format(11): Cell(None),
                        'C{}'.format(11): Cell(None),
                        }

        self.assertEqual(make_part_names_list(sheet = sheet_with_3_pos), ["A720000_K00_Pos0582_3xNorm_3xGes_Mat_3_3547_t=3_3mm", "A720000_K00-01_Pos0202_1xNorm_Mat_3_3547"])

        print()


if __name__=="__main__":
    unittest.main()