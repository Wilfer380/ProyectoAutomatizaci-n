from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from models.asset_record import AssetRecord
from services.process_service import ProcessService


class ProcessServiceBlockPlanningTests(unittest.TestCase):
    def test_plan_blocks_splits_28_records_into_two_blocks(self) -> None:
        records = [self._record(index) for index in range(1, 29)]

        blocks = ProcessService._plan_blocks(records)

        self.assertEqual(2, len(blocks))
        self.assertEqual(27, len(blocks[0]))
        self.assertEqual(1, len(blocks[1]))
        self.assertEqual([record.asset_id for record in records[:27]], [record.asset_id for record in blocks[0]])
        self.assertEqual([record.asset_id for record in records[27:]], [record.asset_id for record in blocks[1]])

    def test_plan_blocks_keeps_27_records_in_a_single_block(self) -> None:
        records = [self._record(index) for index in range(1, 28)]

        blocks = ProcessService._plan_blocks(records)

        self.assertEqual(1, len(blocks))
        self.assertEqual(27, len(blocks[0]))
        self.assertEqual([record.asset_id for record in records], [record.asset_id for record in blocks[0]])

    @staticmethod
    def _record(index: int) -> AssetRecord:
        return AssetRecord(
            row_index=index,
            asset_id=f"A{index:03d}",
            asset_name=f"Asset {index}",
            section="SAP",
        )


if __name__ == "__main__":
    unittest.main()
