def test_workflow_uses_total_pl_lineage_helper():
    from src import total_pl_updater, workflow

    assert (
        workflow.candidate_has_business_lineage
        is total_pl_updater.candidate_has_business_lineage
    )
