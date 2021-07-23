import ObtainResults


class Switcher(object):
    def func_to_call(self, argument, SapModel):
        """Dispatch method"""
        method_name = 'func_' + str(argument)
        # Get the method from 'self'. Default to a lambda.
        method = getattr(self, method_name)
        # Call the method as we return it
        return method(SapModel)

    def func_0(self, SapModel):
        result = ObtainResults.get_planDimension(SapModel)
        return result

    def func_1(self, SapModel):
        result = ObtainResults.get_modulusOfElasticity(SapModel)
        return result

    def func_2(self, SapModel):
        result = ObtainResults.get_structuralPlanDensityIndex(SapModel)
        return result

    def func_3(self, SapModel):
        result = ObtainResults.get_massStory(SapModel)
        return result

    def func_4(self, SapModel):
        result = ObtainResults.get_DynamicProperty(SapModel)
        return result

    def func_5(self, SapModel):
        result = ObtainResults.get_JointResults(SapModel)
        return result

    def func_6(self, SapModel):
        result = ObtainResults.get_lateralDeflectionIndex(SapModel)
        return result

    def func_7(self, SapModel):
        result = ObtainResults.get_shorteningIndex(SapModel)
        return result

    def func_8(self, SapModel):
        result = ObtainResults.get_BRI(SapModel)
        return result

    def func_9(self, SapModel):
        result = ObtainResults.get_axialCapacityIndex(SapModel)
        return result

    def func_10(self, SapModel):
        result = ObtainResults.get_instabilitySafetyFactor(SapModel)
        return result

    def func_11(self, SapModel):
        result = ObtainResults.get_overtuningSafetyFactor(SapModel)
        return result

    def func_12(self, SapModel):
        result = ObtainResults.get_buckingFactor(SapModel)
        return result

    def func_13(self, SapModel):
        result = ObtainResults.get_perimeterIndexes(SapModel)
        return result

    def func_14(self, SapModel):
        result = ObtainResults.get_storyMassIrregularity(SapModel)
        return result

    def func_15(self, SapModel):
        result = ObtainResults.get_storyStiffnessIrregularity(SapModel)
        return result

    def func_16(self, SapModel):
        result = ObtainResults.get_torsionalIrregularityIndex(SapModel)
        return result

    def func_17(self, SapModel):
        result = ObtainResults.get_InherentTorsionRatio(SapModel)
        return result

    def func_18(self, SapModel):
        result = ObtainResults.get_vibrationSeriviceabilityIndex(SapModel)
        return result

    def func_19(self, SapModel):
        result = ObtainResults.get_baseShear(SapModel)
        return result

    def func_20(self, SapModel):
        result = ObtainResults.get_carbonEmission(SapModel)
        return result

    def func_21(self, SapModel):
        result = ObtainResults.gust_effect_factor(SapModel)
        return result

