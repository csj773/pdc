from __future__ import annotations

import build_final_deliverables as b


def keyed(flights: list[b.Flight]) -> dict[tuple, b.Flight]:
    return {b.flight_identity(flight): flight for flight in flights}


def field_tuple(flight: b.Flight) -> tuple:
    return (
        flight.date,
        flight.flt,
        flight.dep,
        flight.arr,
        flight.reg,
        flight.dc,
        flight.off,
        flight.on,
        round(flight.blk_min),
        flight.takeoff,
        flight.landing,
        round(flight.act_min),
        round(flight.ngt_min),
        round(flight.ifr_min),
        flight.to_count,
        flight.ldg_count,
        flight.autoland,
        flight.toga,
        round(flight.pic_min),
        flight.remarks,
    )


def main() -> None:
    pilotlog_flights = b.read_pilotlog_export_flights()
    authoritative_flights = (
        b.read_filled_logbook_flights(b.SYNCED_AUTHORITATIVE_XLSX)
        if b.SYNCED_AUTHORITATIVE_XLSX.exists()
        else []
    )
    if not pilotlog_flights:
        raise SystemExit(f"No usable flights found in {b.PILOTLOG_EXPORT_XLSX}")

    pilotlog_by_key = keyed(pilotlog_flights)
    authoritative_by_key = keyed(authoritative_flights)

    added = len(set(pilotlog_by_key) - set(authoritative_by_key))
    deleted = len(set(authoritative_by_key) - set(pilotlog_by_key))
    modified = sum(
        1
        for key in set(pilotlog_by_key) & set(authoritative_by_key)
        if field_tuple(pilotlog_by_key[key]) != field_tuple(authoritative_by_key[key])
    )

    pilotlog_flights.sort(key=lambda f: (f.date, f.off or b.time(0, 0), f.flt))
    totals = b.compute_totals(b.page_chunks(pilotlog_flights))

    original_output = b.XLSX_OUT
    b.XLSX_OUT = b.SYNCED_AUTHORITATIVE_XLSX
    try:
        b.write_xlsx(pilotlog_flights, totals)
    finally:
        b.XLSX_OUT = original_output

    print(f"SYNC_ADDED {added}")
    print(f"SYNC_DELETED {deleted}")
    print(f"SYNC_MODIFIED {modified}")
    print(f"SYNC_FLIGHTS {len(pilotlog_flights)}")
    print(f"SYNCED_AUTHORITATIVE_XLSX {b.SYNCED_AUTHORITATIVE_XLSX}")


if __name__ == "__main__":
    main()
