"""
利回り計算・ボリュームチェック計算モジュール
"""
from math import floor
from schemas import VolumeCheckInput, VolumeCheckResult


def calc_rooms(land_area: float, far: float, living_ratio: float, room_area: float) -> int:
    """部屋数自動計算"""
    floor_area = land_area * far / 100
    rooms = floor(floor_area * living_ratio / 100 / room_area)
    return max(rooms, 1)


def calc_volume(input: VolumeCheckInput) -> VolumeCheckResult:
    """ボリュームチェック計算"""
    land_price = input.land_price          # 万円
    land_area = input.land_area            # ㎡
    far = input.far                        # %
    monthly_rent = input.monthly_rent      # 万円/部屋
    room_area = input.room_area            # ㎡
    build_cost = input.build_cost_per_tsubo  # 万円/坪
    living_ratio = input.living_ratio      # %
    demolition_unit = input.demolition_unit  # 万円/㎡

    # 延床面積
    floor_area = land_area * far / 100

    # 部屋数
    if input.rooms:
        rooms = input.rooms
    else:
        rooms = calc_rooms(land_area, far, living_ratio, room_area)

    # 年間賃料 (万円)
    annual_rent = monthly_rent * rooms * 12

    # 建設費用 (万円) - 坪単価から計算
    construction = floor_area / 3.3 * build_cost

    # 解体費 (万円)
    demolition = 0.0
    if input.has_old_house:
        # 既存建物の床面積（推定）= 土地面積 × 建蔽率 × 2階分
        estimated_old_area = land_area * input.bcr / 100 * 2
        demolition = estimated_old_area * demolition_unit

    # 諸費用 (万円) - 空欄で自動
    if input.misc_cost is not None:
        misc = input.misc_cost
    else:
        misc = land_price * 0.07 + construction * 0.04

    # 総投資額 (万円)
    total = land_price + construction + demolition + misc

    # 利回り (%)
    yield_pct = (annual_rent / total * 100) if total > 0 else 0.0

    # 坪単価 (万円/坪)
    tsubo_price = (land_price / (land_area / 3.3)) if land_area > 0 else 0.0

    # 一種単価 (万円/坪): 容積率で割り戻した実質坪単価
    ichi_tsubo = (tsubo_price / (far / 100)) if far > 0 else 0.0

    return VolumeCheckResult(
        rooms=rooms,
        floor_area=round(floor_area, 1),
        annual_rent=round(annual_rent, 1),
        construction=round(construction, 1),
        demolition=round(demolition, 1),
        misc=round(misc, 1),
        total_investment=round(total, 1),
        yield_pct=round(yield_pct, 2),
        tsubo_price=round(tsubo_price, 1),
        ichi_tsubo_price=round(ichi_tsubo, 1),
    )


def calc_yield_for_property(
    land_price: float,
    land_area: float,
    far: float = 200.0,
    monthly_rent: float = 5.0,
    room_area: float = 20.0,
    build_cost_per_tsubo: float = 80.0,
    living_ratio: float = 70.0,
    demolition_unit: float = 3.0,
    has_old_house: bool = False,
) -> float:
    """物件の利回り計算（スカラー版）"""
    inp = VolumeCheckInput(
        land_price=land_price,
        land_area=land_area,
        far=far,
        monthly_rent=monthly_rent,
        room_area=room_area,
        build_cost_per_tsubo=build_cost_per_tsubo,
        living_ratio=living_ratio,
        demolition_unit=demolition_unit,
        has_old_house=has_old_house,
    )
    result = calc_volume(inp)
    return result.yield_pct


def yield_color_class(yield_pct: float) -> str:
    """利回りに応じたCSSクラス"""
    if yield_pct >= 8.0:
        return "text-green-600 font-bold"
    elif yield_pct >= 7.0:
        return "text-yellow-600 font-bold"
    else:
        return "text-red-500"
