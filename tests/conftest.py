"""공용 pytest 픽스처."""

import pytest

from core.spec import load_spec_master
from scripts import make_samples


@pytest.fixture(scope="session")
def spec_master(tmp_path_factory):
    """make_samples와 같은 내용의 spec_master.xlsx를 임시 폴더에 만들어 읽은 SpecMaster."""
    path = tmp_path_factory.mktemp("spec") / "spec_master.xlsx"
    make_samples.write_spec_master(path)
    return load_spec_master(path)
