import unittest

from partname_module.partname import PartName


class Cell():

    def __init__(self, value):
        self.value = value

    def __repr__(self):
        return '{}'.format(self.value)

class TestPartName(unittest.TestCase):

    row1 = 1
    sheet1 = {
            'G1': Cell('720000'),
            'A{}'.format(row1): Cell('K00'),
            'B{}'.format(row1): Cell('0582'),
            'D{}'.format(row1): Cell('3'),
            'E{}'.format(row1): Cell('3'),
            'H{}'.format(row1): Cell('3.3547'),
            'L{}'.format(row1): Cell('T=3,3'),
            'C{}'.format(row1): Cell('A')
            }

    row2 = 32
    sheet2 = {
            'G1': Cell('720510'),
            'A{}'.format(row2): Cell('K00-01'),
            'B{}'.format(row2): Cell('0202'),
            'D{}'.format(row2): Cell('1'),
            'E{}'.format(row2): Cell(''),
            'H{}'.format(row2): Cell('3.3547'),
            'L{}'.format(row2): Cell(''),
            'C{}'.format(row2): Cell('A')
            }

    def test_make_a_word_order(self):
        partname1 = PartName(row_number=self.row1, sheet=self.sheet1)
        partname2 = PartName(row_number=self.row2, sheet=self.sheet2)

        self.assertEqual(partname1.make_a_word_order(), "A720000_K00_Pos0582_3xNorm_3xGes_Mat_3_3547_t=3_3mm")

        self.assertEqual(partname2.make_a_word_order(), "A720510_K00-01_Pos0202_1xNorm_Mat_3_3547")

if __name__=="__main__":
    unittest.main()