from bge_pipeline.statistics import compare_train_test, dataset_statistics


def rows(answer_lengths):
    return [
        {"question_id": f"q{i}", "document_id": f"d{i}", "question": "س" * (i + 2), "content_text": "پ" * length}
        for i, length in enumerate(answer_lengths)
    ]


def test_similar_distributions_pass_and_shifted_fail():
    train = dataset_statistics(rows([100, 200, 300, 400, 500]))
    same = dataset_statistics(rows([100, 200, 300, 400, 500]))
    shifted = dataset_statistics(rows([300, 600, 900, 1200, 1500]))
    assert compare_train_test(train, same)["passed"]
    assert not compare_train_test(train, shifted)["passed"]
