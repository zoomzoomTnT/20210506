from unittest import TestCase

from Agent import Agent


class TestAgent(TestCase):

    def test_from_file(self):
        lst = Agent.from_file('resources/20210506.xlsx', '代理人')
        for a in lst:
            assert type(a) is Agent
