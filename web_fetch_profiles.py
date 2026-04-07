# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class WebFetchProfile:
    profile_id: str
    base_url: str
    source_xpath: str
    date_xpath: str
    date_prev_xpath: str
    date_next_xpath: str
    pre_click_xpath: str
    login_input_xpath: str
    login_password_xpath: str
    login_confirm_xpath: str


def little_champion_profile() -> WebFetchProfile:
    return WebFetchProfile(
        profile_id="little_champion_home",
        base_url="https://tainan-production-little-champion.pre-stage.cc/#/home/",
        source_xpath='//*[@id="meal-calc"]/div/div/div[2]/table[2]',
        date_xpath='//*[@id="meal-calc"]/div/div/div[1]/span[1]',
        date_prev_xpath='//*[@id="meal-calc"]/div/div/div[1]/span[2]/img',
        date_next_xpath='//*[@id="meal-calc"]/div/div/div[1]/span[3]/img',
        pre_click_xpath='//*[@id="app"]/div/div/div/div/div[1]/div[2]/button[3]',
        login_input_xpath='//*[@id="app"]/div/div/div[2]/input',
        # 目前網站畫面未見密碼輸入框，先保留欄位供後續 profile 擴充。
        login_password_xpath="",
        login_confirm_xpath='//*[@id="app"]/div/div/div[2]/button',
    )

