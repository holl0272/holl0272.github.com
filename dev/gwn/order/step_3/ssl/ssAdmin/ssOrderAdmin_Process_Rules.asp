<%
Dim maryProcessingActions
Dim maryProcessingRules

'********************************************************************************************************************************************************************************************************
'
' Order Processing Actions
' Note Processing actions must be ordered by priority

mlngRuleCounter = -1

'Order Action Rules
Dim enProcessingAction_None: enProcessingAction_None = updateProcessingArraySize(maryProcessingActions, incrementRule)
	maryProcessingActions(enProcessingAction_None) =					Array(mlngRuleCounter, _
																			  "Take no action", _
																			  0, _
																			  False _
																			  )
Dim enProcessingAction_DeleteOrder: enProcessingAction_DeleteOrder = updateProcessingArraySize(maryProcessingActions, incrementRule)
	maryProcessingActions(enProcessingAction_DeleteOrder) =				Array(mlngRuleCounter, _
																			  "Delete order", _
																			  0, _
																			  False _
																			  )
Dim enProcessingAction_SetOrderToFlagged: enProcessingAction_SetOrderToFlagged = updateProcessingArraySize(maryProcessingActions, incrementRule)
	maryProcessingActions(enProcessingAction_SetOrderToFlagged) =	Array(mlngRuleCounter, _
																			  "Update internal order status to requires review", _
																			  0, _
																			  True _
																			  )
Dim enProcessingAction_ChangeInternalStatusToFlagForManualProcessing: enProcessingAction_ChangeInternalStatusToFlagForManualProcessing = updateProcessingArraySize(maryProcessingActions, incrementRule)
	maryProcessingActions(enProcessingAction_ChangeInternalStatusToFlagForManualProcessing) =			Array(mlngRuleCounter, _
																			  "Update internal flag to requires manual processing", _
																			  0, _
																			  True _
																			  )
Dim enProcessingAction_EmailSpotlight: enProcessingAction_EmailSpotlight = updateProcessingArraySize(maryProcessingActions, incrementRule)
	maryProcessingActions(enProcessingAction_EmailSpotlight) =			Array(mlngRuleCounter, _
																			  "Email order to Manufacturer Spotlight", _
																			  0, _
																			  True _
																			  )
Dim enProcessingAction_Email_PO_GenericVendor: enProcessingAction_Email_PO_GenericVendor = updateProcessingArraySize(maryProcessingActions, incrementRule)
	maryProcessingActions(enProcessingAction_Email_PO_GenericVendor) =			Array(mlngRuleCounter, _
																			  "Email PO to Vendor;Update internal flag to ordered with vendor", _
																			  0, _
																			  False _
																			  )
Dim enProcessingAction_EDI_CWR: enProcessingAction_EDI_CWR = updateProcessingArraySize(maryProcessingActions, incrementRule)
	maryProcessingActions(enProcessingAction_EDI_CWR) =			Array(mlngRuleCounter, _
																			  "EDI to CWR;Update internal flag to ordered with vendor", _
																			  0, _
																			  False _
																			  )

'********************************************************************************************************************************************************************************************************
'
' Order Processing Rules
' Note order of processing rules is of no importance EXCEPT last one for fraud scoring check must be after all rules which add to fraud potential

mlngRuleCounter = -1
Dim enProcessingRule_OldOrder: enProcessingRule_OldOrder = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_OldOrder) =				Array(mlngRuleCounter, _
																		  "Order older than 60 days", _
																		  enProcessingAction_None, _
																		  0 _
																		  )
Dim enProcessingRule_NoOrderDetails: enProcessingRule_NoOrderDetails = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_NoOrderDetails) =			Array(mlngRuleCounter, _
																		  "Orders with no order details", _
																		  enProcessingAction_None, _
																		  0 _
																		  )

'********************************************************************************************************************************************************************************************************
'Fraud Scoring Rules
'International check
Dim enProcessingRule_InternationalShipping: enProcessingRule_InternationalShipping = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_InternationalShipping) =	Array(mlngRuleCounter, _
																		  "Orders with International Shipping Addresses", _
																		  enProcessingAction_ChangeInternalStatusToFlagForManualProcessing, _
																		  7 _
																		  )
Dim enProcessingRule_InternationalBilling: enProcessingRule_InternationalBilling = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_InternationalBilling) =	Array(mlngRuleCounter, _
																		  "Orders with International Billing Addresses", _
																		  enProcessingAction_ChangeInternalStatusToFlagForManualProcessing, _
																		  7 _
																		  )
Dim enProcessingRule_AddressesDoNotMatch: enProcessingRule_AddressesDoNotMatch = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_AddressesDoNotMatch) =		Array(mlngRuleCounter, _
																		  "Orders where Shipping/Billing Addresses do not match", _
																		  enProcessingAction_None, _
																		  0 _
																		  )
Dim enProcessingRule_HighDollarOrder: enProcessingRule_HighDollarOrder = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_HighDollarOrder) =			Array(mlngRuleCounter, _
																		  "High Dollar Order: Order subtotal over $" & cdblOrderThresholdRisk, _
																		  enProcessingAction_ChangeInternalStatusToFlagForManualProcessing, _
																		  5 _
																		  )
Dim enProcessingRule_VeryHighDollarOrder: enProcessingRule_VeryHighDollarOrder = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_VeryHighDollarOrder) =			Array(mlngRuleCounter, _
																		  "Very High Dollar Order: Order subtotal over $" & cdblOrderThresholdAbsolute, _
																		  enProcessingAction_ChangeInternalStatusToFlagForManualProcessing, _
																		  10 _
																		  )
Dim enProcessingRule_ScoreCVV: enProcessingRule_ScoreCVV = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_ScoreCVV) =				Array(mlngRuleCounter, _
																		  "Score Fraud Potential: CCV result", _
																		  enProcessingAction_None, _
																		  0 _
																		  )
Dim enProcessingRule_ScoreAVS: enProcessingRule_ScoreAVS = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_ScoreAVS) =				Array(mlngRuleCounter, _
																		  "Score Fraud Potential: AVS result", _
																		  enProcessingAction_None, _
																		  0 _
																		  )
Dim enProcessingRule_ScorePriorCustomer: enProcessingRule_ScorePriorCustomer = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_ScorePriorCustomer) =		Array(mlngRuleCounter, _
																		  "Score Fraud Potential: Is prior customer", _
																		  enProcessingAction_None, _
																		  -5 _
																		  )
'Check Fraud Score MUST be accomplished after fraud ratings complete
Dim enProcessingRule_CheckFraudScore: enProcessingRule_CheckFraudScore = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_CheckFraudScore) =			Array(mlngRuleCounter, _
																		  "Flag high risk orders. Fraud score over " & cdblFraudThreshold, _
																		  enProcessingAction_ChangeInternalStatusToFlagForManualProcessing, _
																		  cdblFraudThreshold _
																		  )

'********************************************************************************************************************************************************************************************************
'Now begin the mfg/vend specific processing
Dim enProcessingRule_PaidByPO: enProcessingRule_PaidByPO = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_PaidByPO) =			Array(mlngRuleCounter, _
																		  "Paid By PO", _
																		  enProcessingAction_ChangeInternalStatusToFlagForManualProcessing, _
																		  0 _
																		  )
Dim enProcessingRule_PaidByPayPal: enProcessingRule_PaidByPayPal = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_PaidByPayPal) =			Array(mlngRuleCounter, _
																		  "Paid By PayPal", _
																		  enProcessingAction_ChangeInternalStatusToFlagForManualProcessing, _
																		  0 _
																		  )
Dim enProcessingRule_Mfg_Spotlight: enProcessingRule_Mfg_Spotlight = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_Mfg_Spotlight) =			Array(mlngRuleCounter, _
																		  "Order Item from Spotlight", _
																		  enProcessingAction_EmailSpotlight, _
																		  0 _
																		  )
Dim enProcessingRule_Vend_GenericVendor: enProcessingRule_Vend_GenericVendor = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_Vend_GenericVendor) =			Array(mlngRuleCounter, _
																		  "Order Item from Vendor: Generic Vendor", _
																		  enProcessingAction_Email_PO_GenericVendor, _
																		  0 _
																		  )
Dim enProcessingRule_Vend_CWR: enProcessingRule_Vend_CWR = updateProcessingArraySize(maryProcessingRules, incrementRule)
	maryProcessingRules(enProcessingRule_Vend_CWR) =			Array(mlngRuleCounter, _
																		  "Order Item from Vendor: CWR", _
																		  enProcessingAction_EDI_CWR, _
																		  0 _
																		  )
'Payment Type: PO Go to require review
'All internationals must go to review
'Fraud screen ; cease further processing; check setting to see if status is "approved for processing"
'Clean out processing log over X days

'Call writeRules
%>


