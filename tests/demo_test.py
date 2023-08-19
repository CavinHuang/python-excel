from demo import version
import unittest

class Test_demo(unittest.TestCase):
    def test_version(self):
        self.assertEqual(version,"0.0.1")

if __name__ == '__main__':
    unittest.main()