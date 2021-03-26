from mappings import get_color_maps, get_sub_slots

class Slot():
    def __init__(self, slot_num, course_number, is_tut, is_ta):
        self.slot_num = slot_num
        self.course_number = course_number
        self.is_tut = is_tut
        self.is_ta = is_ta

        slots = get_sub_slots()
        if self.slot_num in slots.keys():
            self.is_sub_slot = False
            self.color = get_color_maps(slots[slot_num][0])
        else:
            self.is_sub_slot = True
            self.color = get_color_maps(self.slot_num)
