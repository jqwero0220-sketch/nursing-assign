async def process_student_row_async(row: pd.Series) -> StudentInput:
    mode_raw = str(row.get('이동수단', '대중교통')).strip()
    station_info = str(row.get('인근역', ''))
    address = str(row.get('주소', ''))
    default_time = int(row['소요시간_분'])

    travel_time, transfers, walk_time = get_cached_route_info(address, station_info, mode_raw, default_time)
    mfi = calculate_fatigue_index(travel_time, transfers, walk_time)

    # 💡 이동수단 정밀 분류 로직 수정
    if '버스' in mode_raw and ('전철' in mode_raw or '지하철' in mode_raw):
        detail_mode = '지하철+버스'
    elif '버스' in mode_raw:
        detail_mode = '시내/시외버스'
    elif '전철' in mode_raw or '지하철' in mode_raw:
        detail_mode = '지하철(전철)'
    else:
        detail_mode = mode_raw if mode_raw != '' else '대중교통'  # 엑셀에 적힌 텍스트 그대로 표시

    return StudentInput(
        student_id=str(row['학번']),
        name=str(row['이름']),
        gender=str(row['성별']),
        gpa=float(row['GPA']),
        birth_year=int(row['출생연도']),
        address=address,
        nearest_station=station_info,
        travel_time_minutes=travel_time,
        transfers=transfers,
        walk_time_minutes=walk_time,
        transit_mode=detail_mode,
        fatigue_index=mfi
    )
