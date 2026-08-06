from tools.ad347_data import build_ad347_record


def create_ad347(name, department, work_date):
    return build_ad347_record(
        name=name,
        department=department,
        work_date=work_date,
    )
